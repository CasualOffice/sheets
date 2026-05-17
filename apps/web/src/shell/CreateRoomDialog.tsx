import { useRef, useState } from 'react';
import { Icon } from './Icon';
import { useUniverAPI } from '../use-univer';
import { useWorkbook } from '../use-workbook';
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
      // Both are best-effort: a missing snapshot just means joiners
      // fall back to the xlsx path, which still works.
      if (api) {
        const wb = api.getActiveWorkbook();
        try {
          const blob = await exportCurrentWorkbookAsXlsxBlob(api);
          if (blob) {
            const form = new FormData();
            form.append('file', blob, 'seed.xlsx');
            const seedRes = await fetch(`/api/rooms/${body.roomId}/seed`, {
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
        try {
          if (wb && typeof CompressionStream !== 'undefined') {
            // Walk the in-memory snapshot — wb.save() is a deep clone
            // but we'd pay the same cost in the xlsx exporter, so
            // this is essentially free.
            const snapshot = wb.save();
            const json = JSON.stringify(snapshot);
            const gzipped = await gzipString(json);
            const snapRes = await fetch(`/api/rooms/${body.roomId}/snapshot`, {
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
          }
        } catch (err) {
          console.warn('[share-room] failed to upload snapshot cache', err);
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

  const openRoom = () => {
    if (!created) return;
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
    // Mark this tab as the room's owner. CollabDriver reads this on the
    // destination /r/<id> page and skips the seed download — the owner
    // already has the workbook in memory.
    try {
      sessionStorage.setItem(`casual.collab.owner.${created.roomId}`, '1');
    } catch {
      /* falls through — owner just re-imports their own xlsx, slow but correct */
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
                onClick={openRoom}
              >
                Open the room
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
      <span className="share-room__label" style={emphasis ? { color: 'var(--color-accent)' } : undefined}>
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
