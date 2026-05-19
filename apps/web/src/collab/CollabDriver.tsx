import { useCallback, useEffect, useMemo, useRef, useState, type ReactNode } from 'react';
import * as Y from 'yjs';
import { HocuspocusProvider } from '@hocuspocus/provider';
import { useUniverAPI } from '../use-univer';
import { useWorkbook } from '../use-workbook';
import { useLoading } from '../loading-context';
import { xlsxToWorkbookData } from '../xlsx';
import { startBridge, type BridgeHandle } from './bridge';
import { CollabContext, type CollabRole, type CollabStatus } from './collab-context';
import { useCharts } from '../charts/charts-context';
import type { ChartModel } from '../charts/types';
import { PresenceContext } from './presence-context';
import {
  colorForName,
  getDisplayName,
  setDisplayName,
  suggestAnonName,
  wasNamePrompted,
  markNamePrompted,
  type Identity,
} from './presence';
import { usePresenceWire } from './usePresenceWire';
import { NamePrompt } from './NamePrompt';
import { PresenceLayer } from './PresenceLayer';

/**
 * Owns the co-edit join flow:
 *   1. Parse roomId + role from the URL (`/r/:id?role=view|write`).
 *   2. Pre-flight `/api/rooms/:id/info` to know if a password prompt is
 *      needed. Skipped on builds without the server (Pages demo) —
 *      those just show the self-host banner.
 *   3. Prompt for a password if required; pass it on the WebSocket URL.
 *   4. Spin up the Yjs doc + HocuspocusProvider + bridge.
 *   5. Wire presence/awareness (name, color, selection) and render the
 *      remote-cursor overlay.
 *   6. Publish transport status + role + presence through context so the
 *      rest of the shell (indicator, view-only banner, avatar stack) can
 *      read it.
 */
export function CollabDriver({ children }: { children?: ReactNode }) {
  const api = useUniverAPI();
  const workbook = useWorkbook();
  const loading = useLoading();
  const charts = useCharts();
  const handleRef = useRef<BridgeHandle | null>(null);
  const docRef = useRef<Y.Doc | null>(null);
  // Cleanup for the charts ↔ Yjs bridge wired up when a doc connects.
  const chartsSyncDisposeRef = useRef<(() => void) | null>(null);

  const [provider, setProvider] = useState<HocuspocusProvider | null>(null);
  const [needsSelfHost, setNeedsSelfHost] = useState(false);
  const [roomId, setRoomId] = useState<string | null>(null);
  const [role, setRole] = useState<CollabRole>('write');
  const [status, setStatus] = useState<CollabStatus>('off');
  // Password prompt state: needs to render an input the user can submit.
  // `null` = no prompt; an object = open with this room id pending.
  const [passwordPrompt, setPasswordPrompt] = useState<{ roomId: string; role: CollabRole; error?: string } | null>(null);
  // The actual password (kept in memory only) once the user has provided
  // one — needed to retry on reconnect after a transient drop.
  const passwordRef = useRef<string>('');

  // Local identity. `null` until we've decided we're joining a room.
  const [identity, setIdentity] = useState<Identity | null>(null);
  const [needsNamePrompt, setNeedsNamePrompt] = useState(false);

  // Effect 1: discover the room (URL) and decide whether to prompt or join.
  useEffect(() => {
    if (!api) return;
    const id = readRoomFromLocation();
    if (!id) return;

    if (!isCollabEnabled()) {
      setNeedsSelfHost(true);
      console.info(
        '[collab] /r/%s requested but VITE_COLLAB_ENABLED is not set — self-host with Docker to enable',
        id,
      );
      return;
    }

    const requestedRole: CollabRole = readRoleFromLocation();
    setRoomId(id);
    setRole(requestedRole);

    // Resolve display name. Already set? Use it. Never prompted? Open
    // the prompt with a generated suggestion. Prompted-and-dismissed?
    // Fall through with a generated name (no second prompt).
    const stored = getDisplayName();
    if (stored) {
      setIdentity({ name: stored, color: colorForName(stored) });
    } else {
      const suggestion = suggestAnonName();
      setIdentity({ name: suggestion, color: colorForName(suggestion) });
      if (!wasNamePrompted()) setNeedsNamePrompt(true);
    }

    let cancelled = false;
    (async () => {
      const info = await fetchRoomInfo(id);
      if (cancelled) return;

      // Load the room's starting workbook before bringing the bridge
      // up — otherwise we'd ship a blank grid to peers and overwrite
      // any local edits made between mount and seed-apply. The owner
      // is marked with a sessionStorage flag when they navigated here,
      // so they skip the round-trip (they already *have* the workbook
      // in memory — replacing it would just churn).
      if ((info?.hasSnapshot || info?.hasSeed) && !wasOwnerOfRoom(id)) {
        let loadFailed = false;
        try {
          setStatus('connecting');
          // Show the loading overlay for the joiner. We don't know the
          // file name yet (the server only ships bytes), so use the
          // room id as a stand-in label.
          loading.set({ fileName: `room ${id}`, phase: 'reading' });

          // Fast path: try the pre-parsed gzipped snapshot first. This
          // skips ExcelJS entirely — multi-second win on big workbooks.
          // Browser handles `content-encoding: gzip` transparently, so
          // we just read the JSON.
          let data: import('@univerjs/core').IWorkbookData | null = null;
          let snapshotAttempted = false;
          if (info?.hasSnapshot && typeof DecompressionStream !== 'undefined') {
            snapshotAttempted = true;
            try {
              const res = await fetch(`/api/rooms/${encodeURIComponent(id)}/snapshot`);
              if (res.ok) {
                loading.set({ phase: 'mounting' });
                data = (await res.json()) as import('@univerjs/core').IWorkbookData;
              } else if (!info?.hasSeed) {
                // No seed fallback — propagate the HTTP error.
                throw new Error(`Server returned ${res.status} fetching room snapshot.`);
              } else {
                console.warn('[collab] snapshot fast-path returned', res.status, '— falling back to xlsx');
              }
            } catch (err) {
              if (!info?.hasSeed) throw err;
              console.warn('[collab] snapshot fast-path failed, falling back to xlsx', err);
            }
          }

          // Slow path: parse the xlsx in the worker.
          if (!data && info?.hasSeed) {
            const res = await fetch(`/api/rooms/${encodeURIComponent(id)}/seed`);
            if (!res.ok) {
              throw new Error(`Server returned ${res.status} fetching room seed.`);
            }
            const buf = await res.arrayBuffer();
            loading.set({ phase: 'parsing', sizeBytes: buf.byteLength });
            data = await xlsxToWorkbookData(buf);
            loading.set({ phase: 'mounting' });
          }

          if (cancelled) return;
          if (data) {
            workbook.replaceWorkbook(data, 'xlsx');
            // Univer's unit-swap is async — wait a frame so the new
            // unit is wired into the facade before the bridge attaches.
            await new Promise((r) => requestAnimationFrame(() => r(null)));
          } else if (snapshotAttempted || info?.hasSeed) {
            throw new Error('Room contents could not be loaded (no snapshot or seed returned).');
          }
        } catch (err) {
          loadFailed = true;
          const msg = err instanceof Error ? err.message : String(err);
          console.error('[collab] failed to apply seed', err);
          loading.set({
            fileName: `room ${id}`,
            phase: 'reading',
            error: `Couldn't load this room: ${msg}`,
          });
        } finally {
          // Only auto-dismiss when the load succeeded. If it failed, the
          // overlay stays open in error mode so the user can read why
          // before clicking Dismiss.
          if (!loadFailed) requestAnimationFrame(() => loading.set(null));
        }
        if (cancelled || loadFailed) return;
      }

      if (!info) {
        // Room not found — fall through to the connect attempt anyway;
        // the server will close the upgrade with a clean error.
        join(id, requestedRole, '');
        return;
      }
      if (info.needsPassword) {
        // The owner just stashed their fresh password before navigating —
        // skip the prompt if we have it. sessionStorage scope dies with
        // the tab, so this never leaks across sharing boundaries.
        const stashed = readStashedPassword(id);
        if (stashed) {
          join(id, requestedRole, stashed);
        } else {
          setPasswordPrompt({ roomId: id, role: requestedRole });
        }
      } else {
        join(id, requestedRole, '');
      }
    })();

    return () => {
      cancelled = true;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [api]);

  // Effect 2: tear down on unmount. The join() helper does its own
  // teardown when it's called again (e.g. reconnect with new password).
  useEffect(
    () => () => {
      teardown();
    },
    [],
  );

  const join = useCallback((id: string, joinRole: CollabRole, password: string): void => {
    if (!api) return;
    teardown();
    passwordRef.current = password;
    setStatus('connecting');

    const baseWs = wsUrl();
    const sep = baseWs.includes('?') ? '&' : '?';
    // Hocuspocus reads the document name from its own handshake; we put
    // the same name in the query string so the server's upgrade handler
    // can validate the password against the right room BEFORE the
    // protocol handshake completes.
    const url = `${baseWs}${sep}room=${encodeURIComponent(id)}${password ? `&p=${encodeURIComponent(password)}` : ''}`;

    const doc = new Y.Doc();
    const next = new HocuspocusProvider({ url, name: id, document: doc });
    const handle = startBridge(api, doc, {
      role: joinRole,
      awareness: next.awareness ?? undefined,
      // When a peer's compaction snapshot lands, replace our local
      // workbook with it — same path File→Open uses. Without this,
      // late joiners + restored sessions miss the workbook state.
      onSnapshotReceived: (wb) => {
        workbook.replaceWorkbook(wb, 'xlsx');
      },
    });

    const onStatus = (ev: { status: string }) => {
      if (ev.status === 'connected') setStatus('live');
      else if (ev.status === 'connecting') setStatus('connecting');
      else setStatus('offline');
    };
    next.on('status', onStatus);
    // 4xx close codes from our upgrade handler surface as the provider
    // never moving past `connecting`. Listen for the underlying close
    // so we can flip to `denied` when auth fails.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (next as any).on?.('close', (ev: { code?: number }) => {
      if (ev?.code === 401 || ev?.code === 4401) {
        setStatus('denied');
        setPasswordPrompt({
          roomId: id,
          role: joinRole,
          error: 'Incorrect password — try again.',
        });
      }
    });

    docRef.current = doc;
    handleRef.current = handle;
    setProvider(next);

    // Charts ride a dedicated Y.Map keyed by chart id. The bridge's
    // op-log only carries Univer mutations (cell writes, structural
    // edits); our chart state lives in React, so we sync it
    // out-of-band. Single map + observe means inserts/updates/removes
    // converge in one round-trip just like cell edits.
    chartsSyncDisposeRef.current = wireChartsSync(doc, charts, joinRole);

    console.info('[collab] joined room', id, 'as', joinRole);
  }, [api, charts]);

  const teardown = (): void => {
    chartsSyncDisposeRef.current?.();
    chartsSyncDisposeRef.current = null;
    handleRef.current?.dispose();
    setProvider((p) => {
      p?.destroy();
      return null;
    });
    docRef.current?.destroy();
    handleRef.current = null;
    docRef.current = null;
  };

  // Presence wire — runs in this effect so the same instance feeds both
  // the avatar stack and the overlay through context.
  const { peers } = usePresenceWire(api, provider, identity);

  const collabCtx = useMemo(
    () => ({ enabled: isCollabEnabled(), roomId, status, role }),
    [roomId, status, role],
  );

  const presenceCtx = useMemo(
    () => ({
      me: identity,
      peers,
      needsNamePrompt,
      setName: (name: string) => {
        const clean = name.trim();
        const next = clean.length > 0 ? clean : suggestAnonName();
        setDisplayName(next);
        setIdentity({ name: next, color: colorForName(next) });
        setNeedsNamePrompt(false);
      },
      dismissNamePrompt: () => {
        markNamePrompted();
        setNeedsNamePrompt(false);
      },
    }),
    [identity, peers, needsNamePrompt],
  );

  return (
    <CollabContext.Provider value={collabCtx}>
      <PresenceContext.Provider value={presenceCtx}>
        {needsSelfHost && <SelfHostBanner />}
        {role === 'view' && roomId && status === 'live' && <ViewOnlyBanner />}
        {passwordPrompt && (
          <PasswordPrompt
            state={passwordPrompt}
            onCancel={() => {
              setPasswordPrompt(null);
              setStatus('off');
            }}
            onSubmit={(pw) => {
              setPasswordPrompt(null);
              join(passwordPrompt.roomId, passwordPrompt.role, pw);
            }}
          />
        )}
        {needsNamePrompt && identity && (
          <NamePrompt
            suggestion={identity.name}
            onSubmit={(name) => presenceCtx.setName(name)}
            onCancel={() => presenceCtx.dismissNamePrompt()}
          />
        )}
        {roomId && status !== 'off' && <PresenceLayer />}
        {children}
      </PresenceContext.Provider>
    </CollabContext.Provider>
  );
}

/**
 * Notice shown when someone opens a `/r/:roomId` URL on a build that
 * doesn't ship co-editing (i.e. the public GitHub Pages demo). Points
 * them at the self-host instructions on the apex site.
 */
function SelfHostBanner() {
  const [dismissed, setDismissed] = useState(false);
  if (dismissed) return null;
  return (
    <div className="collab-banner" data-testid="collab-banner" role="status">
      <div className="collab-banner__body">
        <strong>Co-editing requires self-hosting.</strong>{' '}
        The hosted demo at <code>sheet.schnsrw.live</code> is single-user. Run
        Casual Sheets with Docker to get rooms — <a href="https://schnsrw.live/#work" rel="noopener">how to self-host →</a>
      </div>
      <button
        type="button"
        className="collab-banner__close"
        aria-label="Dismiss"
        onClick={() => setDismissed(true)}
      >
        ×
      </button>
    </div>
  );
}

/** Persistent banner shown to view-role joiners — explains why their
 *  edits aren't propagating. */
function ViewOnlyBanner() {
  return (
    <div className="collab-banner collab-banner--neutral" data-testid="view-only-banner" role="status">
      <div className="collab-banner__body">
        <strong>View only.</strong>{' '}
        You're joined as a viewer — your edits stay local and don't sync to others.
        Ask the owner for the edit link if you need to change the sheet.
      </div>
    </div>
  );
}

/** Password prompt modal. Render-blocking — user can't get into the
 *  room without satisfying it. */
function PasswordPrompt({
  state,
  onSubmit,
  onCancel,
}: {
  state: { roomId: string; role: CollabRole; error?: string };
  onSubmit: (pw: string) => void;
  onCancel: () => void;
}) {
  const [value, setValue] = useState('');
  return (
    <div className="dialog-backdrop" data-testid="collab-password-backdrop">
      <div
        className="dialog dialog--narrow"
        role="dialog"
        aria-modal="true"
        aria-labelledby="collab-password-title"
        data-testid="collab-password-dialog"
      >
        <div className="dialog__header">
          <h2 className="dialog__title" id="collab-password-title">
            Password required
          </h2>
        </div>
        <div className="dialog__body">
          <p style={{ margin: '0 0 12px', fontSize: 14, lineHeight: 1.5 }}>
            Room <code>{state.roomId}</code> is password-protected.
            Enter the password the owner shared with you.
          </p>
          <input
            autoFocus
            type="password"
            className="page-setup__select"
            data-testid="collab-password-input"
            value={value}
            onChange={(e) => setValue(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Enter' && value.length > 0) onSubmit(value);
            }}
          />
          {state.error && (
            <p
              data-testid="collab-password-error"
              style={{ margin: '10px 0 0', color: '#d93025', fontSize: 13 }}
            >
              {state.error}
            </p>
          )}
        </div>
        <div className="dialog__footer">
          <button
            type="button"
            className="btn-secondary"
            data-testid="collab-password-cancel"
            onClick={onCancel}
          >
            Cancel
          </button>
          <button
            type="button"
            className="btn-primary"
            data-testid="collab-password-submit"
            disabled={value.length === 0}
            onClick={() => onSubmit(value)}
          >
            Join
          </button>
        </div>
      </div>
    </div>
  );
}

async function fetchRoomInfo(
  id: string,
): Promise<{ needsPassword: boolean; hasSeed: boolean; hasSnapshot: boolean; clients: number } | null> {
  try {
    const res = await fetch(`/api/rooms/${encodeURIComponent(id)}/info`);
    if (!res.ok) return null;
    return (await res.json()) as {
      needsPassword: boolean;
      hasSeed: boolean;
      hasSnapshot: boolean;
      clients: number;
    };
  } catch {
    return null;
  }
}

function isCollabEnabled(): boolean {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  if (typeof (window as any).__COLLAB_WS_URL__ === 'string') return true;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const flag = (import.meta.env as any).VITE_COLLAB_ENABLED as string | undefined;
  return flag === '1' || flag === 'true';
}

function readStashedPassword(roomId: string): string | null {
  try {
    return sessionStorage.getItem(`casual.collab.pw.${roomId}`);
  } catch {
    return null;
  }
}

/** True when this tab created the room and navigated here from the share
 *  dialog. The owner's workbook is already loaded — re-fetching the seed
 *  would just churn a multi-second xlsx parse for no benefit. */
function wasOwnerOfRoom(roomId: string): boolean {
  try {
    return sessionStorage.getItem(`casual.collab.owner.${roomId}`) === '1';
  } catch {
    return false;
  }
}

function readRoomFromLocation(): string | null {
  const path = window.location.pathname.match(/^\/r\/([\w-]{4,})\/?$/);
  if (path) return path[1];
  const params = new URLSearchParams(window.location.search);
  const q = params.get('room');
  return q && q.length >= 4 ? q : null;
}

function readRoleFromLocation(): CollabRole {
  const params = new URLSearchParams(window.location.search);
  return params.get('role') === 'view' ? 'view' : 'write';
}

function wsUrl(): string {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const winOverride = (window as any).__COLLAB_WS_URL__ as string | undefined;
  if (typeof winOverride === 'string' && winOverride) return winOverride;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const envOverride = (import.meta.env as any).VITE_COLLAB_WS_URL as string | undefined;
  if (envOverride) return envOverride;
  const proto = window.location.protocol === 'https:' ? 'wss:' : 'ws:';
  return `${proto}//${window.location.host}/yjs`;
}

/**
 * Wire the ChartsContext to a Yjs map so chart edits flow between
 * peers in real time. One Y.Map keyed by chart id:
 *
 *   - Local insert / update / remove (via `__subscribeLocal`) → diff
 *     against the current map and apply add/set/delete in a single
 *     transaction. We diff rather than nuke-and-rewrite so concurrent
 *     edits on different charts don't clobber each other.
 *   - Remote map changes (via `observe`) → snapshot the map into a
 *     ChartModel[] and call `__replaceAll({ fromCollab: true })` so
 *     ChartsContext applies the remote state WITHOUT echoing it back
 *     through `__subscribeLocal` (avoiding ping-pong).
 *
 * View-role clients only consume remote updates — local edits never
 * write to the map, matching the read-only contract the rest of the
 * collab path uses.
 */
function wireChartsSync(
  doc: Y.Doc,
  charts: ReturnType<typeof useCharts>,
  role: CollabRole,
): () => void {
  const map = doc.getMap<ChartModel>('casual-charts');

  // Hydrate from the map immediately. If a peer compacted a snapshot
  // before us, the map may already hold the room's charts.
  if (map.size > 0) {
    charts.__replaceAll(Array.from(map.values()), { fromCollab: true });
  } else if (role === 'write' && charts.charts.length > 0) {
    // We're joining with local charts (e.g. the room owner uploaded a
    // workbook that already contained charts via xlsx) — seed the map.
    doc.transact(() => {
      for (const c of charts.charts) map.set(c.id, c);
    });
  }

  let applyingRemote = false;

  const onLocal = (next: ChartModel[]) => {
    if (role === 'view') return;
    if (applyingRemote) return;
    doc.transact(() => {
      const nextIds = new Set(next.map((c) => c.id));
      // Drop charts that are no longer local.
      for (const id of map.keys()) {
        if (!nextIds.has(id)) map.delete(id);
      }
      // Add / overwrite charts that are. Yjs only encodes a delta when
      // the value actually changed (it serialises new bytes either way,
      // but downstream subscribers diff before re-applying).
      for (const c of next) {
        const cur = map.get(c.id);
        if (!cur || !shallowEqualChart(cur, c)) map.set(c.id, c);
      }
    });
  };
  const unsubLocal = charts.__subscribeLocal(onLocal);

  const onRemote = (event: Y.YMapEvent<ChartModel>, tx: Y.Transaction) => {
    // Ignore our own writes — Y dispatches them too. The tx's origin
    // is the doc itself for `doc.transact(...)`; we mark applyingRemote
    // around `__replaceAll` to suppress the echo below.
    void event;
    void tx;
    applyingRemote = true;
    try {
      charts.__replaceAll(Array.from(map.values()), { fromCollab: true });
    } finally {
      applyingRemote = false;
    }
  };
  map.observe(onRemote);

  return () => {
    unsubLocal();
    map.unobserve(onRemote);
  };
}

function shallowEqualChart(a: ChartModel, b: ChartModel): boolean {
  // Cheap structural compare to avoid an extra Yjs encode when the
  // chart object reference changed but the values are identical
  // (React's `setCharts` always returns a new array on every set).
  return JSON.stringify(a) === JSON.stringify(b);
}
