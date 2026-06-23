import { useCallback, useEffect, useMemo, useRef, useState, type ReactNode } from 'react';
import * as Y from 'yjs';
import type { HocuspocusProvider } from '@hocuspocus/provider';
import { useUniverAPI } from '../use-univer';
import { useWorkbook } from '../use-workbook';
import { useLoading } from '../loading-context';
import { useToast } from '../shell/toast/toast-context';
import { xlsxToWorkbookData } from '../xlsx';
// Phase 3 (thin host): the app no longer hand-rolls the Yjs doc + provider +
// bridge — it calls the SDK's one-shot `attachCollab` and layers its own host
// concerns (presence, charts, password re-prompt, divergence) on the returned
// provider/doc.
import {
  attachCollab,
  type BridgeHandle,
  type CollabHandle,
  type ReplayFailureRecord,
} from '@casualoffice/sheets/collab';
import {
  CollabContext,
  type CollabRole,
  type CollabStatus,
  type SyncHealth,
} from './collab-context';
import { viteEnv, windowStringGlobal } from '../univer-facade';
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
import { applyViewOnlyMode } from './view-mode';
import { applyCommentOnly } from '@casualoffice/sheets/sheets';
import { useCurrentUser } from '../auth/auth-context';
import { parseShareMeta, shareMetaUrl, sharePasswordKey, type ShareMeta } from './share-meta';

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
  const toast = useToast();
  const charts = useCharts();
  // Personal-mode pre-fill — UX_AUDIT.md §4.4 / Phase 4 #17. A user
  // who has already identified themselves to the host shouldn't be
  // asked to do it again to join a collab room on their own box.
  // Returns null when not signed in (Mode 1 / Mode 2 / WOPI / anon
  // share) — the existing whimsical fallback takes over.
  const currentUser = useCurrentUser();
  // `collabRef` owns the SDK collab session (doc + provider + bridge +
  // detach). `handleRef` keeps pointing at the bridge so the existing
  // replay-failure subscriptions below stay unchanged.
  const collabRef = useRef<CollabHandle | null>(null);
  const handleRef = useRef<BridgeHandle | null>(null);
  const docRef = useRef<Y.Doc | null>(null);
  // Cleanup for the charts ↔ Yjs bridge wired up when a doc connects.
  const chartsSyncDisposeRef = useRef<(() => void) | null>(null);
  // Cleanup for the view-only permission gate. Applied after join for
  // `role=view` joiners; re-applied on every workbook swap (snapshot
  // replace, late-join seed apply) since the permission is per-unit-id
  // and a swap creates a fresh unit.
  const viewModeDisposeRef = useRef<(() => void) | null>(null);

  const [provider, setProvider] = useState<HocuspocusProvider | null>(null);
  // Mirror docRef into state so the CollabContext value updates when
  // the active room's doc changes. HistoryPanel and any future
  // op-log consumer subscribe via context and rerender on transitions.
  const [doc, setDoc] = useState<Y.Doc | null>(null);
  const [needsSelfHost, setNeedsSelfHost] = useState(false);
  // Set when a `?share=<token>` link resolves to invalid/expired via the
  // public /meta probe — we render a clear dead-link state and never
  // attempt the WS connect (sharing-model §6.1 join side).
  const [shareInvalid, setShareInvalid] = useState(false);
  const [roomId, setRoomId] = useState<string | null>(null);
  const [role, setRole] = useState<CollabRole>('write');
  const [status, setStatus] = useState<CollabStatus>('off');
  // Password prompt state: needs to render an input the user can submit.
  // `null` = no prompt; an object = open with this room id pending.
  // `kind` distinguishes the anonymous-room `?p=` gate from the secure
  // share-link `?sp=` gate — they're DISTINCT passwords and the
  // re-prompt-on-reject path routes the submitted value differently.
  const [passwordPrompt, setPasswordPrompt] = useState<{
    roomId: string;
    role: CollabRole;
    kind: 'room' | 'share';
    error?: string;
  } | null>(null);
  // The actual password (kept in memory only) once the user has provided
  // one — needed to retry on reconnect after a transient drop.
  const passwordRef = useRef<string>('');
  // Secure share-link capability read off the page URL once (sharing-model
  // §6.1). `null` for anonymous rooms. Stable for the page's lifetime, so a
  // ref read at join time (incl. reconnect re-joins) is correct and keeps
  // `join`'s signature unchanged.
  const shareRef = useRef<{ share: string; sp?: string } | null>(readShareFromLocation());
  // Pending password-prompt resolver. The connect flow awaits this so
  // we can hold the seed/snapshot fetch off until the user submits.
  // Replaced wholesale by each `promptForPassword` call; the previous
  // promise is left dangling — that's fine, no work is gated on it.
  const passwordPromptResolverRef = useRef<{
    resolve: (pw: string) => void;
    reject: (err: Error) => void;
  } | null>(null);
  const promptForPassword = useCallback(
    (
      id: string,
      joinRole: CollabRole,
      kind: 'room' | 'share' = 'room',
      error?: string,
    ): Promise<string> => {
      return new Promise<string>((resolve, reject) => {
        passwordPromptResolverRef.current = { resolve, reject };
        setPasswordPrompt({ roomId: id, role: joinRole, kind, error });
      });
    },
    [],
  );

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

    // Resolve display name. Precedence:
    //   1. Stored localStorage name (the user has previously joined a
    //      room and set / accepted a name — keep it stable).
    //   2. Personal-mode signed-in username (UX_AUDIT.md §4.4). We
    //      ALSO write it through `setDisplayName` so the same identity
    //      survives a sign-out → anonymous-join round trip.
    //   3. Whimsical anon name + open the prompt.
    const stored = getDisplayName();
    if (stored) {
      setIdentity({ name: stored, color: colorForName(stored) });
    } else if (currentUser) {
      const name = currentUser.username;
      setDisplayName(name);
      setIdentity({ name, color: colorForName(name) });
    } else {
      const suggestion = suggestAnonName();
      setIdentity({ name: suggestion, color: colorForName(suggestion) });
      if (!wasNamePrompted()) setNeedsNamePrompt(true);
    }

    let cancelled = false;
    (async () => {
      // ── STEP 0 — secure share-link pre-flight (sharing-model §6.1). ──
      // When the page URL carries `?share=<token>`, probe the PUBLIC
      // /meta endpoint BEFORE touching the room: the token IS the
      // capability, so we can discover (a) whether it's dead and (b)
      // whether a `?sp=` password is needed, without connecting.
      const share = shareRef.current;
      if (share && !share.sp) {
        const meta = await fetchShareMeta(share.share);
        if (cancelled) return;
        if (meta && !meta.valid) {
          // Dead link — surface a clear state, do NOT connect.
          setShareInvalid(true);
          setStatus('off');
          return;
        }
        // `meta === null` means the probe itself failed (server
        // unreachable / older build without /meta). Fall through and let
        // the WS connect decide — a password-gated token will be rejected
        // there and routed to the share re-prompt below, matching the
        // wrong-password path.
        if (meta && meta.valid && meta.hasPassword) {
          const stashed = readStashedSharePassword(share.share);
          if (stashed) {
            shareRef.current = { ...share, sp: stashed };
          } else {
            let sp = '';
            try {
              sp = await promptForPassword(id, requestedRole, 'share');
            } catch {
              // User cancelled the share-password prompt — leave idle.
              setStatus('off');
              return;
            }
            if (cancelled) return;
            stashSharePassword(share.share, sp);
            shareRef.current = { ...share, sp };
          }
        }
      }

      const info = await fetchRoomInfo(id);
      if (cancelled) return;

      if (!info) {
        // Room not found — fall through to the connect attempt anyway;
        // the server will close the upgrade with a clean error.
        join(id, requestedRole, '');
        return;
      }

      // STEP 1 — resolve the password BEFORE any content fetch. Earlier
      // versions downloaded the workbook first and then prompted, which
      // (a) leaked content to anyone who knew the room URL and (b) let
      // the joiner see fully-rendered cells behind the password modal
      // for a second or two. Hold here until the user submits a
      // password we can pass to the seed/snapshot fetches.
      let password = '';
      if (info.needsPassword) {
        const stashed = readStashedPassword(id);
        if (stashed) {
          password = stashed;
        } else {
          // Show the prompt and wait for submit. Resolve to the
          // entered password; cancellation throws so the outer
          // try/finally tears down cleanly.
          try {
            password = await promptForPassword(id, requestedRole);
          } catch {
            // User cancelled — bail out, leave the page idle.
            setStatus('off');
            return;
          }
          if (cancelled) return;
        }
      }

      // STEP 2 — fetch the room's starting workbook. We always do this
      // (even for the owner) because the owner navigates via
      // `window.location.href = ...` to /r/<id>, which is a full page
      // load — the in-memory workbook from the share dialog is gone by
      // the time we get here. The earlier "owner already has the data
      // in memory" optimisation only worked when navigation was a
      // React-router push, which it isn't.
      if (info.hasSnapshot || info.hasSeed) {
        let loadFailed = false;
        try {
          setStatus('connecting');
          loading.set({ fileName: `room ${id}`, phase: 'reading' });

          // Pass the password on the fetch as a header — the server uses
          // it to gate /seed and /snapshot the same way it gates the WS
          // upgrade. The header is more correct than `?p=` because it
          // doesn't end up in access logs / browser history.
          const authHeader: Record<string, string> = password
            ? { 'x-room-password': password }
            : {};

          let data: import('@univerjs/core').IWorkbookData | null = null;
          let snapshotAttempted = false;
          if (info.hasSnapshot && typeof DecompressionStream !== 'undefined') {
            snapshotAttempted = true;
            try {
              const res = await fetch(`/api/rooms/${encodeURIComponent(id)}/snapshot`, {
                headers: authHeader,
              });
              if (res.ok) {
                loading.set({ phase: 'mounting' });
                data = (await res.json()) as import('@univerjs/core').IWorkbookData;
              } else if (res.status === 401) {
                throw new Error('Password rejected by server.');
              } else if (!info.hasSeed) {
                throw new Error(`Server returned ${res.status} fetching room snapshot.`);
              } else {
                console.warn(
                  '[collab] snapshot fast-path returned',
                  res.status,
                  '— falling back to xlsx',
                );
              }
            } catch (err) {
              if (!info.hasSeed) throw err;
              console.warn('[collab] snapshot fast-path failed, falling back to xlsx', err);
            }
          }

          if (!data && info.hasSeed) {
            const res = await fetch(`/api/rooms/${encodeURIComponent(id)}/seed`, {
              headers: authHeader,
            });
            if (res.status === 401) {
              throw new Error('Password rejected by server.');
            }
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
          } else if (snapshotAttempted || info.hasSeed) {
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
          if (!loadFailed) requestAnimationFrame(() => loading.set(null));
        }
        if (cancelled || loadFailed) return;
      }

      // STEP 3 — bring up the Yjs connection. By this point either
      // there's no password OR we already have it.
      join(id, requestedRole, password);
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

  const join = useCallback(
    (id: string, joinRole: CollabRole, password: string): void => {
      if (!api) return;
      teardown();
      passwordRef.current = password;
      setStatus('connecting');

      // Hand the room to the SDK. `attachCollab` builds the WS URL (room +
      // password + role onto the base — the server's upgrade handler validates
      // the password against the right room and flips view joiners to read-only
      // BEFORE the protocol handshake), creates the Yjs doc + Hocuspocus
      // provider (10 s reconnect timeout so a drop surfaces as `offline` fast),
      // and starts the mutation bridge. We layer presence / charts / close
      // handling / divergence onto the returned provider + doc below.
      // When a secure share token is on the page URL, forward it (+ optional
      // join password) on the WS upgrade and let the SERVER decide the role —
      // it resolves the token to a role bound to this room at mint time and
      // ignores any client `?role=`. So we pass `share` to the SDK (which then
      // omits `role=` from the URL) but still pass `joinRole` for the LOCAL
      // belt-and-braces view-only gate below; the server stays authoritative.
      const share = shareRef.current ?? undefined;
      const collab = attachCollab(api, {
        room: id,
        server: wsUrl(),
        password: password || undefined,
        // The bridge only knows view|write. A `comment` joiner broadcasts like a
        // writer (so their comment mutations sync) — the applyCommentOnly veto
        // already stops cell mutations from firing, so none reach the bridge.
        role: joinRole === 'comment' ? 'write' : joinRole,
        share,
        // Map the SDK's coarse status onto our richer CollabStatus (which also
        // carries 'off' / 'denied', driven elsewhere in this component).
        onStatus: (s) => setStatus(s),
        // When a peer's compaction snapshot lands, replace our local workbook —
        // same path File→Open uses. The bridge AWAITS this so its next replay
        // rewrites unitId against the NEW workbook (docs/COLLAB-FIXES.md issue
        // 2). The swap creates a fresh unit id, so re-apply the view-only
        // permission against it (the prior permission point is now stale).
        onSnapshot: async (wb) => {
          workbook.replaceWorkbook(wb, 'xlsx');
          await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));
          if ((joinRole === 'view' || joinRole === 'comment') && api) {
            const swapped = api.getActiveWorkbook();
            if (swapped) {
              viewModeDisposeRef.current?.();
              viewModeDisposeRef.current =
                joinRole === 'comment'
                  ? applyCommentOnly(api)
                  : applyViewOnlyMode(api, swapped.getId());
            }
          }
        },
      });
      const doc = collab.doc;
      const next = collab.provider;
      const handle = collab.bridge;

      // Auth-failure routing. The server now completes the WS upgrade and
      // then closes the socket with code 4401 when the password is wrong
      // (see apps/server/src/yjs.ts) — that way the browser surfaces it
      // as a real CloseEvent we can switch on, rather than a 1006 that
      // looks indistinguishable from a network drop.
      //
      // The forwarded close event from Hocuspocus is a plain CloseEvent;
      // `event` is the underlying browser CloseEvent.
      const handleDeniedClose = () => {
        // Tear down THIS provider so the underlying provider's auto-reconnect
        // doesn't keep hammering the server with the wrong password.
        teardown();
        setStatus('denied');
        // A share-token join was rejected by the server. The most common
        // cause is a wrong (or missing) `?sp=` share password — the server
        // collapses every token failure to one closed connection, but for a
        // token we know carried a password the actionable case is "wrong
        // password". Re-prompt for the SHARE password (`?sp=`), drop the
        // stale stash, and retry forwarding the new value via shareRef.
        const activeShare = shareRef.current;
        if (activeShare) {
          forgetSharePassword(activeShare.share);
          void promptForPassword(id, joinRole, 'share', 'Incorrect password — try again.').then(
            (sp) => {
              stashSharePassword(activeShare.share, sp);
              shareRef.current = { ...activeShare, sp };
              join(id, joinRole, passwordRef.current);
            },
            () => setStatus('off'),
          );
          return;
        }
        // Anonymous-room path — re-prompt for the room `?p=` password. On
        // submit, the resolver fires `join(id, joinRole, pw)` again.
        void promptForPassword(id, joinRole, 'room', 'Incorrect password — try again.').then(
          (pw) => join(id, joinRole, pw),
          () => setStatus('off'),
        );
      };
      // Hocuspocus `close` payload is `{ event: CloseEvent }`.
      next.on('close', (payload: { event?: { code?: number } } | undefined) => {
        const code = payload?.event?.code;
        if (code === 4401 || code === 1008) handleDeniedClose();
      });
      // Belt-and-braces: Hocuspocus also emits `authenticationFailed` when
      // its own onAuthenticate hook rejects. We don't use that hook today,
      // but wiring this means a future auth hook just works.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (next as any).on?.('authenticationFailed', handleDeniedClose);

      collabRef.current = collab;
      docRef.current = doc;
      handleRef.current = handle;
      setProvider(next);
      setDoc(doc);
      // Expose for e2e diagnostics / browser-devtools poking. Same
      // policy as __univerAPI — no secrets, but invaluable when a real
      // user reports "the sync looks broken" and we need to inspect
      // the live Yjs document state without rebuilding.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (window as any).__hocuspocusProvider = next;

      // Charts ride a dedicated Y.Map keyed by chart id. The bridge's
      // op-log only carries Univer mutations (cell writes, structural
      // edits); our chart state lives in React, so we sync it
      // out-of-band. Single map + observe means inserts/updates/removes
      // converge in one round-trip just like cell edits.
      chartsSyncDisposeRef.current = wireChartsSync(doc, charts, joinRole);

      // View-only joiners get Univer's WorkbookEditablePermission flipped
      // to false. The editor refuses to open and edit menu items go
      // disabled — the bridge's role-gate (which drops outbound
      // mutations) becomes belt-and-braces rather than the only line of
      // defence. Apply after the first frame so the workbook unit is
      // wired into the facade.
      if ((joinRole === 'view' || joinRole === 'comment') && api) {
        requestAnimationFrame(() => {
          const wb = api.getActiveWorkbook();
          if (!wb) return;
          viewModeDisposeRef.current?.();
          viewModeDisposeRef.current =
            joinRole === 'comment'
              ? applyCommentOnly(api)
              : applyViewOnlyMode(api, wb.getId());
        });
      }

      console.info('[collab] joined room', id, 'as', joinRole);
    },
    [api, charts],
  );

  const teardown = (): void => {
    chartsSyncDisposeRef.current?.();
    chartsSyncDisposeRef.current = null;
    viewModeDisposeRef.current?.();
    viewModeDisposeRef.current = null;
    // One call tears down the bridge + provider + doc (idempotent).
    collabRef.current?.detach();
    collabRef.current = null;
    handleRef.current = null;
    docRef.current = null;
    setProvider(null);
    setDoc(null);
  };

  // Presence wire — runs in this effect so the same instance feeds both
  // the avatar stack and the overlay through context.
  const { peers } = usePresenceWire(api, provider, identity);

  // Divergence detector — peers broadcast their Y.Doc state-vector hex
  // every 5 s via awareness. Compute aggregate sync health: in-sync
  // when every visible peer's `sv` matches our local doc's; syncing
  // when they disagree but the disagreement is recent (< 15 s); and
  // diverged otherwise. Local SV is read directly from our doc; the
  // peers' SV comes off awareness. Recomputed on a 2 s interval to
  // catch the "syncing → diverged" transition without doing work on
  // every awareness change.
  const [syncHealth, setSyncHealth] = useState<SyncHealth>('in-sync');
  const firstDisagreeAtRef = useRef<number>(0);
  useEffect(() => {
    if (status !== 'live' || !docRef.current) {
      setSyncHealth('in-sync');
      firstDisagreeAtRef.current = 0;
      return;
    }
    const compute = () => {
      const doc = docRef.current;
      if (!doc) return;
      let local = '';
      try {
        const sv = Y.encodeStateVector(doc);
        let hex = '';
        for (let i = 0; i < sv.length; i += 1) {
          const b = sv[i];
          if (b < 16) hex += '0';
          hex += b.toString(16);
        }
        local = hex;
      } catch {
        /* doc destroyed mid-tick */
        return;
      }
      // A peer counts only if they reported an `sv` in the last 30 s —
      // older readings are stale (peer might have disconnected, idle
      // tab throttled by browser, etc.) and would false-positive.
      const now = Date.now();
      const fresh = peers.filter(
        (p) =>
          typeof p.sv === 'string' &&
          typeof p.svAt === 'number' &&
          now - (p.svAt as number) < 30_000,
      );
      if (fresh.length === 0) {
        setSyncHealth('in-sync');
        firstDisagreeAtRef.current = 0;
        return;
      }
      const allMatch = fresh.every((p) => p.sv === local);
      if (allMatch) {
        setSyncHealth('in-sync');
        firstDisagreeAtRef.current = 0;
      } else {
        if (firstDisagreeAtRef.current === 0) firstDisagreeAtRef.current = now;
        const elapsed = now - firstDisagreeAtRef.current;
        setSyncHealth(elapsed > 15_000 ? 'diverged' : 'syncing');
      }
    };
    compute();
    const id = setInterval(compute, 2000);
    return () => clearInterval(id);
  }, [status, peers]);

  // Reconnect feedback — fire a transient toast when the WS
  // transitions from `offline` back to `live`. Pre-toast, the
  // OfflineBanner appeared during the outage and then silently
  // disappeared on reconnect, leaving the user unsure whether the
  // queued edits actually flushed (audit finding 1.4). Track the
  // previous status via a ref so the effect fires once per
  // transition rather than on every status read.
  const prevStatusRef = useRef<CollabStatus>('off');
  useEffect(() => {
    const prev = prevStatusRef.current;
    if (prev === 'offline' && status === 'live') {
      toast.success(
        queuedLocal > 0
          ? `Reconnected — ${queuedLocal} ${queuedLocal === 1 ? 'change' : 'changes'} synced`
          : 'Reconnected — sync resumed',
      );
    }
    prevStatusRef.current = status;
    // queuedLocal isn't in deps on purpose: by the time we re-enter
    // live, the queuedLocal effect (further down) resets it to 0;
    // we want the count AT the moment of transition, which the
    // closure captures via the `queuedLocal` reference above. Read
    // it lazily so the toast text reflects what actually queued
    // during the outage.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [status]);

  // Replay-failure counter — the bridge increments this every time
  // a remote mutation throws on local apply. Each one is a candidate
  // divergence (peer's edit didn't land here, so state vectors will
  // disagree). Sticky for the session; refresh-recommended UX when
  // > 0. See v0.1 audit finding #2: failed replays were silently
  // logged to console.warn with no user-facing surface.
  const [replayFailures, setReplayFailures] = useState(0);
  const [replayDeadLetter, setReplayDeadLetter] = useState<readonly ReplayFailureRecord[]>([]);
  useEffect(() => {
    const handle = handleRef.current;
    if (!handle) return;
    setReplayFailures(handle.getReplayFailures());
    setReplayDeadLetter(handle.getReplayDeadLetter());
    const off = handle.subscribeReplayFailures((n) => setReplayFailures(n));
    const offDl = handle.subscribeReplayDeadLetter((entries) => setReplayDeadLetter(entries));
    return () => {
      off();
      offDl();
    };
    // re-subscribe whenever the bridge handle is replaced (room
    // re-join after disconnect, etc.). We key on `doc` since each
    // bridge handle has a fresh doc instance.
  }, [doc]);

  // Queued-mutation counter: how many of OUR mutation records have
  // landed in the op log since we last knew the server was alive.
  // Yjs writes to the local Y.Array immediately even when the WS is
  // offline; the count is what's piled up locally waiting for the
  // next flush. Reset to 0 every time we re-enter `live`. Lets the
  // offline banner / indicator show "3 changes queued" so users
  // don't think their edits vanished. See feedback from the v0.1
  // collab UX audit.
  const [queuedLocal, setQueuedLocal] = useState(0);
  const queuedOffsetRef = useRef<number>(0);
  useEffect(() => {
    if (!doc) {
      setQueuedLocal(0);
      queuedOffsetRef.current = 0;
      return;
    }
    if (status === 'live') {
      setQueuedLocal(0);
      queuedOffsetRef.current = 0;
      return;
    }
    if (status !== 'offline' && status !== 'connecting') return;
    // Use a loose-typed Y.Array since the bridge's MutationRecord
    // shape isn't exported and only the `c` (clientId) field matters
    // here. Mirror the LOG_KEY constant from bridge.ts (`'ops'`).
    type LogRecord = { c?: string };
    const log = doc.getArray<LogRecord>('ops');
    const myId = String(doc.clientID);
    queuedOffsetRef.current = log.length;
    const recount = () => {
      let count = 0;
      for (let i = queuedOffsetRef.current; i < log.length; i += 1) {
        const rec = log.get(i);
        if (rec && String(rec.c) === myId) count += 1;
      }
      setQueuedLocal(count);
    };
    log.observe(recount);
    recount();
    return () => {
      log.unobserve(recount);
    };
  }, [doc, status]);

  const collabCtx = useMemo(
    () => ({
      enabled: isCollabEnabled(),
      roomId,
      status,
      role,
      syncHealth,
      peerCount: peers.length,
      queuedLocal,
      replayFailures,
      replayDeadLetter,
      doc,
    }),
    [
      roomId,
      status,
      role,
      syncHealth,
      peers.length,
      queuedLocal,
      replayFailures,
      replayDeadLetter,
      doc,
    ],
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
        {shareInvalid && <ShareInvalidBanner />}
        {role === 'view' && roomId && status === 'live' && <ViewOnlyBanner />}
        {role === 'comment' && roomId && status === 'live' && <CommentOnlyBanner />}
        {roomId && status === 'offline' && <OfflineBanner queuedLocal={queuedLocal} />}
        {passwordPrompt && (
          <PasswordPrompt
            state={passwordPrompt}
            currentName={identity?.name ?? suggestAnonName()}
            onCancel={() => {
              setPasswordPrompt(null);
              passwordPromptResolverRef.current?.reject(new Error('cancelled'));
              passwordPromptResolverRef.current = null;
            }}
            onSubmit={({ password, name }) => {
              // Apply the (possibly edited) display name BEFORE the WS
              // upgrade resolves — peers see the chosen name from their
              // first awareness broadcast, never the prior identity.
              if (name && name !== identity?.name) {
                presenceCtx.setName(name);
              }
              setPasswordPrompt(null);
              passwordPromptResolverRef.current?.resolve(password);
              passwordPromptResolverRef.current = null;
            }}
          />
        )}
        {/* Suppress the standalone first-time NamePrompt while the join
         *  prompt is open — the join prompt already carries the
         *  display-name field, and stacking two modals would trap the
         *  user behind an unrelated picker. */}
        {needsNamePrompt && !passwordPrompt && identity && (
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
        <strong>Co-editing requires self-hosting.</strong> The hosted demo at{' '}
        <code>sheet.casualoffice.org</code> is single-user. Run Casual Sheets with Docker to get
        rooms —{' '}
        <a href="https://casualoffice.org/#work" rel="noopener">
          how to self-host →
        </a>
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

/** Shown when a `?share=<token>` link resolves to invalid/expired via the
 *  public /meta probe (sharing-model §6.1). Render-blocking would be
 *  overkill — there's nothing to connect to, so a clear banner is enough.
 *  We never attempted the WS, so no content leaked behind it. */
function ShareInvalidBanner() {
  return (
    <div
      className="collab-banner collab-banner--warn"
      data-testid="share-invalid-banner"
      role="status"
    >
      <div className="collab-banner__body">
        <strong>This share link is invalid or has expired.</strong> Ask the owner for a fresh link.
      </div>
    </div>
  );
}

/** Persistent banner shown to view-role joiners — explains why their
 *  edits aren't propagating. */
function ViewOnlyBanner() {
  return (
    <div
      className="collab-banner collab-banner--neutral"
      data-testid="view-only-banner"
      role="status"
    >
      <div className="collab-banner__body">
        <strong>View only.</strong> You're joined as a viewer — your edits stay local and don't sync
        to others. Ask the owner for the edit link if you need to change the sheet.
      </div>
    </div>
  );
}

/** Banner for the `comment` share-role: cells are locked but comments work. */
function CommentOnlyBanner() {
  return (
    <div
      className="collab-banner collab-banner--neutral"
      data-testid="comment-only-banner"
      role="status"
    >
      <div className="collab-banner__body">
        <strong>Comment only.</strong> You can add and reply to comments, but cell edits are locked.
        Ask the owner for the edit link to change the sheet.
      </div>
    </div>
  );
}

/** Banner shown when the WS provider reports an offline status (after
 *  Hocuspocus's internal reconnect heuristics have kicked in). The
 *  provider keeps retrying with exponential backoff in the background;
 *  this banner just makes it visible that a reconnect is pending.
 *
 *  When local mutations have piled up while offline, surface the
 *  count so users see their work isn't being lost — the most common
 *  worry when the indicator turns amber. */
function OfflineBanner({ queuedLocal }: { queuedLocal: number }) {
  const tail =
    queuedLocal > 0
      ? `${queuedLocal} ${queuedLocal === 1 ? 'change is' : 'changes are'} queued locally and will sync when the connection is back.`
      : 'Your edits are queued locally and will sync when the connection is back.';
  return (
    <div
      className="collab-banner collab-banner--warn"
      data-testid="offline-banner"
      data-queued-local={queuedLocal}
      role="status"
      aria-live="polite"
    >
      <div className="collab-banner__body">
        <strong>Waiting to reconnect to server…</strong> {tail}
      </div>
    </div>
  );
}

/** Password prompt modal. Render-blocking — user can't get into the
 *  room without satisfying it. */
/**
 * Render-blocking join modal. Two fields:
 *
 *   1. Display name — pre-filled with the current identity name. Peers
 *      see it the moment we connect, so letting the user override here
 *      (instead of forcing them to use the title-bar NamePill *after*
 *      joining) means a fresh tab can pick a fresh persona for one room
 *      without leaking the previous identity even briefly.
 *   2. Password — the existing gate. Disabled the Join button until a
 *      non-empty password is entered (room actually needs one to let us
 *      through the WS upgrade).
 */
function PasswordPrompt({
  state,
  currentName,
  onSubmit,
  onCancel,
}: {
  state: { roomId: string; role: CollabRole; kind: 'room' | 'share'; error?: string };
  currentName: string;
  onSubmit: (args: { password: string; name: string }) => void;
  onCancel: () => void;
}) {
  const [value, setValue] = useState('');
  const [name, setName] = useState(currentName);
  const submit = () => {
    if (value.length === 0) return;
    onSubmit({ password: value, name: name.trim() || currentName });
  };
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
            Joining room {state.roomId}
          </h2>
        </div>
        <div className="dialog__body">
          <label
            style={{
              display: 'block',
              margin: '0 0 6px',
              fontSize: 13,
              color: 'var(--color-fg-muted, #5f6368)',
            }}
            htmlFor="collab-join-name"
          >
            Joining as
          </label>
          <input
            id="collab-join-name"
            type="text"
            maxLength={32}
            className="page-setup__select"
            data-testid="collab-join-name-input"
            value={name}
            onChange={(e) => setName(e.target.value)}
            style={{ marginBottom: 12 }}
            aria-label="Your display name for this room"
          />
          <p style={{ margin: '0 0 12px', fontSize: 14, lineHeight: 1.5 }}>
            {state.kind === 'share' ? (
              <>This share link is password-protected. Enter the password the owner sent you.</>
            ) : (
              <>
                Room <code>{state.roomId}</code> is password-protected. Enter the password the owner
                shared with you.
              </>
            )}
          </p>
          <input
            autoFocus
            type="password"
            className="page-setup__select"
            data-testid="collab-password-input"
            value={value}
            onChange={(e) => setValue(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Enter' && value.length > 0) submit();
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
            onClick={submit}
          >
            Join
          </button>
        </div>
      </div>
    </div>
  );
}

async function fetchRoomInfo(id: string): Promise<{
  needsPassword: boolean;
  hasSeed: boolean;
  hasSnapshot: boolean;
  clients: number;
} | null> {
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

/**
 * Probe the PUBLIC /meta endpoint to discover whether a share token is
 * live and whether it needs a `?sp=` password — BEFORE opening the WS.
 * The token is the capability, so no auth is sent. Returns the parsed
 * meta, or `null` when the probe itself fails (network error, a 503 in
 * an anonymous-only deploy, an older server without /meta, or a
 * malformed body) — callers treat null as "couldn't determine, let the
 * WS connect decide".
 */
async function fetchShareMeta(token: string): Promise<ShareMeta | null> {
  try {
    const res = await fetch(shareMetaUrl(token));
    if (!res.ok) return null;
    return parseShareMeta(await res.json());
  } catch {
    return null;
  }
}

/** Read a stashed share-link password for this token (sessionStorage,
 *  dies with the tab). Keyed by token so a reconnect within the session
 *  doesn't re-prompt — mirrors the room-password `casual.collab.pw.<id>`
 *  pattern but namespaced under `sp` to keep the two distinct. */
function readStashedSharePassword(token: string): string | null {
  try {
    return sessionStorage.getItem(sharePasswordKey(token));
  } catch {
    return null;
  }
}

function stashSharePassword(token: string, password: string): void {
  if (!password) return;
  try {
    sessionStorage.setItem(sharePasswordKey(token), password);
  } catch {
    /* private mode — joiner will be re-prompted on reconnect, fine */
  }
}

function forgetSharePassword(token: string): void {
  try {
    sessionStorage.removeItem(sharePasswordKey(token));
  } catch {
    /* swallow */
  }
}

function isCollabEnabled(): boolean {
  if (windowStringGlobal('__COLLAB_WS_URL__')) return true;
  const flag = viteEnv('VITE_COLLAB_ENABLED');
  return flag === '1' || flag === 'true';
}

function readStashedPassword(roomId: string): string | null {
  try {
    return sessionStorage.getItem(`casual.collab.pw.${roomId}`);
  } catch {
    return null;
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
  const r = params.get('role');
  if (r === 'view') return 'view';
  if (r === 'comment') return 'comment';
  return 'write';
}

/**
 * Read the secure share-link capability off the PAGE url (sharing-model
 * §6.1). `?share=<token>` is the minted token; `?sp=<password>` is the
 * optional join password for a password-protected token. When a token is
 * present the SERVER is authoritative for the role — the token is forwarded
 * (+ sp) on the WS upgrade and the client `?role=` is ignored. Returns
 * `null` when there's no token, so the anonymous-room path
 * (`?room`/`?p`/`?role`) is completely unchanged.
 */
function readShareFromLocation(): { share: string; sp?: string } | null {
  const params = new URLSearchParams(window.location.search);
  const token = params.get('share');
  if (!token) return null;
  const sp = params.get('sp');
  return sp ? { share: token, sp } : { share: token };
}

function wsUrl(): string {
  const winOverride = windowStringGlobal('__COLLAB_WS_URL__');
  if (winOverride) return winOverride;
  const envOverride = viteEnv('VITE_COLLAB_WS_URL');
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
    if (role !== 'write') return; // view + comment are read-only for charts
    if (applyingRemote) return;
    doc.transact(() => {
      const nextIds = new Set(next.map((c) => c.id));
      // Drop charts that are no longer local.
      for (const id of map.keys()) {
        if (!nextIds.has(id)) map.delete(id);
      }
      // Add / overwrite changed charts. Reference equality is the
      // right diff here: ChartsContext's `update` only creates a new
      // object for the chart that actually changed (it uses
      // `prev.map(c => c.id === id ? {...c, ...patch} : c)` so
      // untouched charts keep their reference). `insert` adds a new
      // ref. `remove` only filters. So `cur !== c` flags ALL real
      // local changes and ignores no-ops in O(1) per chart, vs the
      // previous JSON.stringify-per-chart which scaled with chart
      // payload size and dominated re-render cost on dashboards with
      // many ECharts options.
      for (const c of next) {
        const cur = map.get(c.id);
        if (cur !== c) map.set(c.id, c);
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
