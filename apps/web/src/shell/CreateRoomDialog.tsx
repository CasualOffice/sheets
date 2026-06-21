import { useCallback, useEffect, useRef, useState } from 'react';
import { Icon } from './Icon';
import { useUniverAPI } from '../use-univer';
import { useWorkbook } from '../use-workbook';
import { useFileSource } from '../file-source/context';
import { useCharts } from '../charts/charts-context';
import { exportCurrentWorkbookAsXlsxBlob, loadSpreadsheetFile } from './file-actions';

type Role = 'write' | 'view';

type Stage = 'configure' | 'creating' | 'ready';

type Created = {
  roomId: string;
  writeUrl: string;
  viewUrl: string;
};

/**
 * "Share for co-editing" dialog. Three sections in the configure stage:
 *
 *   1. Optional file picker — if chosen, we load it into the local
 *      workbook *before* creating the room so the owner's bridge has
 *      content to broadcast. If omitted, the current workbook is the
 *      seed (default for "share what I'm already editing").
 *   2. Optional password.
 *   3. Default role — "Anyone with the share link can edit" vs
 *      "Anyone with the share link can view only".
 *
 * After POST /api/rooms succeeds, the dialog flips to the "ready" stage
 * showing two shareable URLs (write + view) with copy buttons. The
 * write/view choice in step 3 just picks which one is highlighted —
 * both URLs are always shown so the owner can hand out either.
 */
export function CreateRoomDialog({ onClose }: { onClose: () => void }) {
  const api = useUniverAPI();
  const workbook = useWorkbook();
  const fileSource = useFileSource();
  const { charts } = useCharts();

  // Secure share-LINK affordance (sharing-model §6.1) is offered only for a
  // PERSONAL-mode saved file (we have a `serverFileId` the /files/:id/shares
  // routes are keyed on + owner-gated by). In Mode 1 / unsaved files we keep
  // today's anonymous-room behaviour untouched — the spoofable `?role=view`
  // URL stays the only sharing surface there.
  const serverFileId = fileSource.kind === 'personal' ? (workbook.meta.serverFileId ?? null) : null;
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [file, setFile] = useState<File | null>(null);
  const [password, setPassword] = useState('');
  const [defaultRole, setDefaultRole] = useState<Role>('write');
  const [stage, setStage] = useState<Stage>('configure');
  const [error, setError] = useState<string | null>(null);
  const [created, setCreated] = useState<Created | null>(null);

  const submit = async () => {
    setError(null);
    setStage('creating');
    try {
      if (file) {
        // Swap local workbook BEFORE allocating the room so the upload
        // below captures the new content.
        await loadSpreadsheetFile(file, api, workbook.replaceWorkbook);
      }
      const res = await fetch('/api/rooms', {
        method: 'POST',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify(password.length > 0 ? { password } : {}),
      });
      if (!res.ok) throw new Error(`server responded ${res.status}`);
      const body = (await res.json()) as { roomId: string };

      // Upload TWO representations of the starting workbook:
      //   1. xlsx bytes — for compatibility and any tooling that wants
      //      the raw file (server stores at /api/rooms/:id/seed).
      //   2. gzipped JSON snapshot — fast-path for joiners so they can
      //      skip the multi-second ExcelJS parse on join.
      //
      // The xlsx upload is *required* for joiners to start from the
      // same workbook, so we await it. The snapshot upload is a pure
      // fast-path — joiners fall back to the xlsx if it's missing, so
      // we fire it in the background. This keeps the dialog's "Ready"
      // state from waiting on two serial worker exports + uploads.
      if (api) {
        const roomId = body.roomId;
        const wb = api.getActiveWorkbook();
        try {
          const blob = await exportCurrentWorkbookAsXlsxBlob(api, { charts });
          if (blob) {
            const form = new FormData();
            form.append('file', blob, 'seed.xlsx');
            const seedRes = await fetch(`/api/rooms/${roomId}/seed`, {
              method: 'POST',
              body: form,
            });
            if (!seedRes.ok) {
              console.warn(
                '[share-room] seed upload failed',
                seedRes.status,
                await seedRes.text().catch(() => ''),
              );
            }
          }
        } catch (err) {
          console.warn('[share-room] failed to serialize xlsx seed', err);
        }
        // Fire-and-forget — joiners fall back to /seed if this is
        // missing or still in flight. Don't block the dialog on it.
        if (wb && typeof CompressionStream !== 'undefined') {
          void (async () => {
            try {
              const snapshot = wb.save();
              const json = JSON.stringify(snapshot);
              const gzipped = await gzipString(json);
              const snapRes = await fetch(`/api/rooms/${roomId}/snapshot`, {
                method: 'POST',
                headers: { 'content-type': 'application/gzip' },
                body: gzipped as BodyInit,
              });
              if (!snapRes.ok) {
                console.warn(
                  '[share-room] snapshot upload failed',
                  snapRes.status,
                  await snapRes.text().catch(() => ''),
                );
              }
            } catch (err) {
              console.warn('[share-room] failed to upload snapshot cache', err);
            }
          })();
        }
      }

      const origin = window.location.origin;
      const writeUrl = `${origin}/r/${body.roomId}`;
      const viewUrl = `${origin}/r/${body.roomId}?role=view`;
      setCreated({ roomId: body.roomId, writeUrl, viewUrl });
      setStage('ready');
    } catch (err) {
      console.error('[share-room] create failed', err);
      setError(
        'Could not start the room. The server may be unreachable — co-editing requires the self-hosted Docker build.',
      );
      setStage('configure');
    }
  };

  const [opening, setOpening] = useState(false);

  const openRoom = async () => {
    if (!created || opening) return;
    setOpening(true);
    // Stash the password in sessionStorage keyed by room id so the owner's
    // own join skips the password prompt — they just typed it. Scoped to
    // sessionStorage so it dies with the tab; never persisted to disk.
    if (password.length > 0) {
      try {
        sessionStorage.setItem(`casual.collab.pw.${created.roomId}`, password);
      } catch {
        /* private mode — owner will be re-prompted, fine */
      }
    }

    // Re-upload the seed + snapshot with the CURRENT workbook state.
    // Between submit (which did the initial upload from a snapshot of
    // the workbook AT CREATE-CLICK time) and now, the owner could have
    // typed into cells — the "Ready" stage doesn't disable editing. If
    // we navigated immediately, those edits would be lost: the page
    // reload nukes in-memory Univer state, and /r/<id> mount downloads
    // the STALE seed the server has from submit time. Re-uploading
    // captures every post-create edit. Best-effort: a network blip
    // here is logged but doesn't block navigation.
    if (api) {
      const auth: Record<string, string> =
        password.length > 0 ? { 'x-room-password': password } : {};
      try {
        const blob = await exportCurrentWorkbookAsXlsxBlob(api, { charts });
        if (blob) {
          const form = new FormData();
          form.append('file', blob, 'seed.xlsx');
          await fetch(`/api/rooms/${created.roomId}/seed`, {
            method: 'POST',
            body: form,
            headers: auth,
          });
        }
        const wb = api.getActiveWorkbook();
        if (wb && typeof CompressionStream !== 'undefined') {
          const snapshot = wb.save();
          const gzipped = await gzipString(JSON.stringify(snapshot));
          await fetch(`/api/rooms/${created.roomId}/snapshot`, {
            method: 'POST',
            headers: { 'content-type': 'application/gzip', ...auth },
            body: gzipped as BodyInit,
          });
        }
      } catch (err) {
        console.warn('[share-room] re-upload before navigate failed', err);
      }
    }

    // Owner always joins as write — the view URL is for the people they
    // share with. Navigate (not assign) so back-button works cleanly.
    window.location.href = created.writeUrl;
  };

  return (
    <div className="dialog-backdrop" data-testid="share-room-backdrop" onClick={onClose}>
      <div
        className="dialog dialog--narrow"
        role="dialog"
        aria-modal="true"
        aria-labelledby="share-room-title"
        data-testid="share-room-dialog"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="dialog__header">
          <h2 className="dialog__title" id="share-room-title">
            {stage === 'ready' ? 'Room ready — share these links' : 'Share for co-editing'}
          </h2>
        </div>

        <div className="dialog__body">
          {stage !== 'ready' && (
            <div className="share-room">
              <div className="share-room__row">
                <label className="share-room__label" htmlFor="share-room-file">
                  Seed with a file (optional)
                </label>
                <div className="share-room__file">
                  <input
                    ref={fileInputRef}
                    id="share-room-file"
                    type="file"
                    accept=".xlsx,.ods,.csv,.tsv,.tab"
                    style={{ display: 'none' }}
                    data-testid="share-room-file-input"
                    onChange={(e) => setFile(e.target.files?.[0] ?? null)}
                  />
                  <button
                    type="button"
                    className="btn-secondary"
                    data-testid="share-room-file-pick"
                    onClick={() => fileInputRef.current?.click()}
                  >
                    <Icon name="folder_open" size="sm" /> Choose file…
                  </button>
                  <span className="share-room__file-name">
                    {file ? file.name : 'Share the current workbook'}
                  </span>
                  {file && (
                    <button
                      type="button"
                      className="btn-secondary"
                      aria-label="Clear chosen file"
                      data-testid="share-room-file-clear"
                      onClick={() => {
                        setFile(null);
                        if (fileInputRef.current) fileInputRef.current.value = '';
                      }}
                    >
                      <Icon name="close" size="sm" />
                    </button>
                  )}
                </div>
              </div>

              <div className="share-room__row">
                <label className="share-room__label" htmlFor="share-room-password">
                  Password (optional)
                </label>
                <input
                  id="share-room-password"
                  type="password"
                  className="page-setup__select"
                  data-testid="share-room-password"
                  placeholder="Leave blank for an open room"
                  value={password}
                  onChange={(e) => setPassword(e.target.value)}
                  autoComplete="new-password"
                />
                <p className="share-room__hint">
                  Anyone with the link still needs the password to join.
                </p>
              </div>

              <div className="share-room__row">
                <span className="share-room__label">Default sharing role</span>
                <div className="share-room__segment" role="radiogroup" aria-label="Default role">
                  <button
                    type="button"
                    role="radio"
                    aria-checked={defaultRole === 'write'}
                    className={
                      'share-room__seg-opt' +
                      (defaultRole === 'write' ? ' share-room__seg-opt--active' : '')
                    }
                    data-testid="share-room-role-write"
                    onClick={() => setDefaultRole('write')}
                  >
                    <span>Edit</span>
                    <span className="share-room__seg-opt-desc">Can change cells & sheets</span>
                  </button>
                  <button
                    type="button"
                    role="radio"
                    aria-checked={defaultRole === 'view'}
                    className={
                      'share-room__seg-opt' +
                      (defaultRole === 'view' ? ' share-room__seg-opt--active' : '')
                    }
                    data-testid="share-room-role-view"
                    onClick={() => setDefaultRole('view')}
                  >
                    <span>View only</span>
                    <span className="share-room__seg-opt-desc">Can read but not edit</span>
                  </button>
                </div>
              </div>

              {error && (
                <p
                  data-testid="share-room-error"
                  style={{ margin: 0, color: '#d93025', fontSize: 13 }}
                >
                  {error}
                </p>
              )}
            </div>
          )}

          {stage === 'ready' && created && (
            <div className="share-room">
              <p className="share-room__hint" style={{ marginTop: 0 }}>
                Both links go to the same room — pick whichever fits the person you're sharing with.
              </p>
              <ShareUrlRow
                label="Edit link"
                description="Anyone with this link can change cells."
                url={created.writeUrl}
                emphasis={defaultRole === 'write'}
                testidPrefix="share-room-write"
              />
              <ShareUrlRow
                label="View-only link"
                description="Anyone with this link can read but not edit."
                url={created.viewUrl}
                emphasis={defaultRole === 'view'}
                testidPrefix="share-room-view"
              />
              {password.length > 0 && (
                <p className="share-room__hint">
                  Send the password separately — it isn't embedded in the link.
                </p>
              )}

              {serverFileId && (
                <ShareLinkSection serverFileId={serverFileId} roomId={created.roomId} />
              )}
            </div>
          )}
        </div>

        <div className="dialog__footer">
          {stage !== 'ready' && (
            <>
              <button
                type="button"
                className="btn-secondary"
                data-testid="share-room-cancel"
                onClick={onClose}
                disabled={stage === 'creating'}
              >
                Cancel
              </button>
              <button
                type="button"
                className="btn-primary"
                data-testid="share-room-create"
                onClick={() => void submit()}
                disabled={stage === 'creating'}
              >
                {stage === 'creating' ? 'Creating room…' : 'Create share link'}
              </button>
            </>
          )}
          {stage === 'ready' && (
            <>
              <button
                type="button"
                className="btn-secondary"
                data-testid="share-room-done"
                onClick={onClose}
              >
                Done
              </button>
              <button
                type="button"
                className="btn-primary"
                data-testid="share-room-open"
                onClick={() => void openRoom()}
                disabled={opening}
              >
                {opening ? 'Opening…' : 'Open the room'}
              </button>
            </>
          )}
        </div>
      </div>
    </div>
  );
}

/**
 * Gzip a JS string via the browser-native CompressionStream. Available
 * in every evergreen browser; we feature-check before calling so an
 * older Safari just skips the snapshot upload and joiners fall back
 * to the xlsx path.
 */
async function gzipString(input: string): Promise<Uint8Array> {
  const enc = new TextEncoder().encode(input);
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const stream = new (window as any).CompressionStream('gzip') as ReadableWritablePair<
    Uint8Array,
    Uint8Array
  >;
  const blob = await new Response(
    new Blob([enc as BlobPart]).stream().pipeThrough(stream as unknown as ReadableWritablePair),
  ).arrayBuffer();
  return new Uint8Array(blob);
}

function ShareUrlRow({
  label,
  description,
  url,
  emphasis,
  testidPrefix,
}: {
  label: string;
  description: string;
  url: string;
  emphasis: boolean;
  testidPrefix: string;
}) {
  const [copied, setCopied] = useState(false);
  const copy = async () => {
    try {
      await navigator.clipboard.writeText(url);
      setCopied(true);
      setTimeout(() => setCopied(false), 1600);
    } catch {
      /* permissions denied — user can still select the field */
    }
  };
  return (
    <div className="share-room__url">
      <span
        className="share-room__label"
        style={emphasis ? { color: 'var(--color-accent)' } : undefined}
      >
        {label}
      </span>
      <span className="share-room__hint">{description}</span>
      <div className="share-room__url-row">
        <input
          type="text"
          readOnly
          className="share-room__url-input"
          value={url}
          data-testid={`${testidPrefix}-url`}
          onFocus={(e) => e.currentTarget.select()}
        />
        <button
          type="button"
          className={'share-room__copy' + (copied ? ' share-room__copy--done' : '')}
          data-testid={`${testidPrefix}-copy`}
          onClick={() => void copy()}
        >
          <Icon name={copied ? 'check' : 'content_copy'} size="sm" />
          {copied ? 'Copied' : 'Copy'}
        </button>
      </div>
    </div>
  );
}

// ── Secure share links (sharing-model §6.1) ──────────────────────────────────

/** A link role exposed in the UI. `comment` is intentionally NOT offered —
 *  the server collapses it to read-only today, so it would mislead. */
type LinkRole = 'view' | 'edit';

/** Public projection returned by GET/POST/PATCH /files/:id/shares — mirrors
 *  `toPublicLink` in apps/server/src/files/personal-shares-routes.ts. The
 *  server's `ShareRole` is broader; we only ever create / show view|edit. */
type ShareLinkDto = {
  token: string;
  roomId: string;
  role: string;
  expiresAt: number | null;
  hasPassword: boolean;
  createdAt: number;
  createdBy: string;
};

const EXPIRY_CHOICES: { label: string; days?: number }[] = [
  { label: 'Never' },
  { label: '7 days', days: 7 },
  { label: '30 days', days: 30 },
];

/**
 * Secure-link minting + management, shown in the "ready" stage for a
 * personal-mode saved file. Unlike the anonymous `?role=view` URL above, a
 * minted link is a SERVER-ENFORCED capability: the token is bound to this
 * room at mint time and the join handshake resolves the role from it (the
 * client `?role=` is ignored). Password is supported by the server but
 * intentionally not exposed in this v1 UI — the join-side password prompt
 * is a separate batch.
 */
function ShareLinkSection({ serverFileId, roomId }: { serverFileId: string; roomId: string }) {
  const [role, setRole] = useState<LinkRole>('view');
  const [expiryIdx, setExpiryIdx] = useState(0);
  const [linkPassword, setLinkPassword] = useState('');
  const [links, setLinks] = useState<ShareLinkDto[]>([]);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const base = `/files/${encodeURIComponent(serverFileId)}/shares`;

  const refresh = useCallback(async () => {
    try {
      const res = await fetch(base, { credentials: 'include' });
      if (!res.ok) throw new Error(`list links: HTTP ${res.status}`);
      const body = (await res.json()) as { links: ShareLinkDto[] };
      // Show ALL of this file's links — the endpoint is already scoped to the
      // file (`/files/:id/shares`). Don't filter by the dialog's anonymous room:
      // the server binds each token to the file's deterministic `pf-<id>` room,
      // so filtering by `created.roomId` hid every minted link. Each row builds
      // its own URL from `link.roomId`.
      setLinks(body.links);
    } catch (err) {
      console.warn('[share-link] list failed', err);
      setError('Could not load existing links.');
    }
  }, [base]);

  useEffect(() => {
    void refresh();
  }, [refresh]);

  const create = async () => {
    setBusy(true);
    setError(null);
    try {
      const expiresInDays = EXPIRY_CHOICES[expiryIdx]?.days;
      const trimmedPw = linkPassword.trim();
      const res = await fetch(`${base}/link`, {
        method: 'POST',
        credentials: 'include',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({
          roomId,
          role,
          ...(expiresInDays ? { expiresInDays } : {}),
          ...(trimmedPw.length > 0 ? { password: trimmedPw } : {}),
        }),
      });
      if (!res.ok) throw new Error(`create link: HTTP ${res.status}`);
      // Clear the password field after a successful mint so a second
      // link isn't accidentally created with the same (now stale) value.
      setLinkPassword('');
      await refresh();
    } catch (err) {
      console.warn('[share-link] create failed', err);
      setError('Could not create the link. Make sure the file is saved and you own it.');
    } finally {
      setBusy(false);
    }
  };

  const revoke = async (token: string) => {
    setError(null);
    try {
      const res = await fetch(`${base}/link/${encodeURIComponent(token)}`, {
        method: 'DELETE',
        credentials: 'include',
      });
      if (!res.ok && res.status !== 204) throw new Error(`revoke link: HTTP ${res.status}`);
      await refresh();
    } catch (err) {
      console.warn('[share-link] revoke failed', err);
      setError('Could not revoke the link.');
    }
  };

  return (
    <div className="share-link" data-testid="share-link-section">
      <div className="share-link__divider" />
      <span className="share-room__label">Secure link (enforced)</span>
      <p className="share-room__hint" style={{ marginTop: 0 }}>
        Unlike the links above, a secure link's view/edit role is enforced by the server — it can't
        be changed by editing the URL. Best for a saved file. Add a password for a second gate —
        send it separately, it isn't part of the link.
      </p>

      <div className="share-link__controls">
        <div className="share-link__field">
          <label className="share-room__label" htmlFor="share-link-role">
            Anyone with this link can
          </label>
          <select
            id="share-link-role"
            className="page-setup__select"
            data-testid="share-link-role"
            value={role}
            onChange={(e) => setRole(e.target.value as LinkRole)}
          >
            <option value="view">View only</option>
            <option value="edit">Edit</option>
          </select>
        </div>
        <div className="share-link__field">
          <label className="share-room__label" htmlFor="share-link-expiry">
            Expires
          </label>
          <select
            id="share-link-expiry"
            className="page-setup__select"
            data-testid="share-link-expiry"
            value={expiryIdx}
            onChange={(e) => setExpiryIdx(Number(e.target.value))}
          >
            {EXPIRY_CHOICES.map((c, i) => (
              <option key={c.label} value={i}>
                {c.label}
              </option>
            ))}
          </select>
        </div>
        <div className="share-link__field">
          <label className="share-room__label" htmlFor="share-link-password">
            Password (optional)
          </label>
          <input
            id="share-link-password"
            type="password"
            className="page-setup__select"
            data-testid="share-link-password"
            placeholder="No password"
            value={linkPassword}
            onChange={(e) => setLinkPassword(e.target.value)}
            autoComplete="new-password"
          />
        </div>
        <button
          type="button"
          className="btn-primary share-link__create"
          data-testid="share-link-create"
          onClick={() => void create()}
          disabled={busy}
        >
          <Icon name="add_link" size="sm" />
          {busy ? 'Creating…' : 'Create link'}
        </button>
      </div>

      {error && (
        <p
          data-testid="share-link-error"
          style={{ margin: '4px 0 0', color: '#d93025', fontSize: 13 }}
        >
          {error}
        </p>
      )}

      <ul className="share-link__list" data-testid="share-link-list">
        {links.length === 0 && (
          <li className="share-link__empty" data-testid="share-link-empty">
            No secure links yet.
          </li>
        )}
        {links.map((link) => (
          <ShareLinkRow key={link.token} link={link} onRevoke={() => void revoke(link.token)} />
        ))}
      </ul>
    </div>
  );
}

/** One row in the secure-link list: a view/edit badge, the copy-able
 *  `?share=` URL, expiry, and Revoke. */
function ShareLinkRow({ link, onRevoke }: { link: ShareLinkDto; onRevoke: () => void }) {
  const [copied, setCopied] = useState(false);
  // The server returns the URL FRAGMENT (`?share=<token>`); the full link is
  // the room URL the host is already on + that fragment. We rebuild it here
  // against the live origin so the copied link works from this deployment.
  const url = `${window.location.origin}/r/${encodeURIComponent(link.roomId)}?share=${encodeURIComponent(link.token)}`;
  const roleLabel = link.role === 'edit' || link.role === 'write' ? 'edit' : 'view';

  const copy = async () => {
    try {
      await navigator.clipboard.writeText(url);
      setCopied(true);
      setTimeout(() => setCopied(false), 1600);
    } catch {
      /* clipboard denied — the input is still selectable */
    }
  };

  return (
    <li className="share-link__item" data-testid="share-link-item" data-role={roleLabel}>
      <div className="share-link__item-head">
        <span className={`share-link__badge share-link__badge--${roleLabel}`}>{roleLabel}</span>
        {link.hasPassword && (
          <span
            className="share-link__lock"
            data-testid="share-link-password-badge"
            title="Password required to join"
          >
            <Icon name="lock" size="sm" />
            password
          </span>
        )}
        <span className="share-room__hint">{formatExpiry(link.expiresAt)}</span>
        <button
          type="button"
          className="share-link__revoke"
          data-testid="share-link-revoke"
          aria-label="Revoke this link"
          onClick={onRevoke}
        >
          <Icon name="delete" size="sm" />
          Revoke
        </button>
      </div>
      <div className="share-room__url-row">
        <input
          type="text"
          readOnly
          className="share-room__url-input"
          value={url}
          data-testid="share-link-url"
          onFocus={(e) => e.currentTarget.select()}
        />
        <button
          type="button"
          className={'share-room__copy' + (copied ? ' share-room__copy--done' : '')}
          data-testid="share-link-copy"
          onClick={() => void copy()}
        >
          <Icon name={copied ? 'check' : 'content_copy'} size="sm" />
          {copied ? 'Copied' : 'Copy'}
        </button>
      </div>
    </li>
  );
}

/** "Never expires" / "Expires <date>" / "Expired" — short, no library. */
function formatExpiry(expiresAt: number | null): string {
  if (expiresAt === null) return 'Never expires';
  if (expiresAt <= Date.now()) return 'Expired';
  const d = new Date(expiresAt);
  return `Expires ${d.toLocaleDateString(undefined, { month: 'short', day: 'numeric', year: 'numeric' })}`;
}
