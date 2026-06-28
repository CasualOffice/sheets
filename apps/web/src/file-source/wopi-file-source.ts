/**
 * Copyright 2026 Casual Office
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

import { xlsxToWorkbookData } from '../xlsx';
import type { FileSource, OpenedWorkbook, RecentEntry, SaveOptions, SaveResult } from './types';

/**
 * Embedded-host `FileSource` (Phase D / Mode 2 of #49). Wraps the
 * server-side WOPI routes (`/wopi/files/:id` + `/contents`) with a
 * single bound file id read from the URL `?access_token=<JWT>` claim
 * (or `?wopi_src=…&access_token=…`, the standard embed shape).
 *
 * Surface mapping
 *
 *   - `listRecent`     — unsupported. The embedding host owns
 *                        navigation; we return a list containing the
 *                        single token-bound file so the home screen
 *                        "recents" strip shows it as a one-tap reopen.
 *   - `subscribeRecent` — single-shot ticker; the list never changes
 *                        within an embedded session.
 *   - `openRecent`     — GET `/wopi/files/:id/contents?access_token=…`.
 *                        Reads `X-WOPI-ItemVersion` for the etag.
 *   - `forgetRecent`   — no-op. Embedded host owns its files.
 *   - `save`           — POST `/wopi/files/:id/contents` with the
 *                        same access token + `X-WOPI-ItemVersion`
 *                        header as If-Match. 409 surfaces as a
 *                        `conflict` SaveResult.
 *
 * Auth flow with Personal mode coexisting on the same server: the
 * cookie isn't sent on WOPI routes (`credentials: 'omit'`) so a
 * personal session can't unlock an embed-only resource. The query
 * `access_token` is the only credential.
 */

type CheckFileInfo = {
  BaseFileName: string;
  Size: number;
  Version: string;
  ReadOnly?: boolean;
  UserCanWrite?: boolean;
};

type WopiBootContext = {
  /** JWT issued by the embedding host. Bound to `fileId`. */
  accessToken: string;
  /** File id from the JWT's `file_id` claim, mirrored in the URL. */
  fileId: string;
};

/** Peek at a JWT payload without verifying — we never trust the
 *  client-decoded shape; the server checks the signature on every
 *  call. The peek lets the source know its bound `file_id` without
 *  a round-trip to `/api/me`. */
function decodeJwtPayload(token: string): Record<string, unknown> | null {
  try {
    const parts = token.split('.');
    if (parts.length !== 3) return null;
    const padded = parts[1].replace(/-/g, '+').replace(/_/g, '/');
    const json = atob(padded + '='.repeat((4 - (padded.length % 4)) % 4));
    return JSON.parse(json) as Record<string, unknown>;
  } catch {
    return null;
  }
}

/** Inspect `window.location` for a WOPI-style access token. Returns
 *  null when none is present — the picker upstream falls back to the
 *  browser source. */
export function detectWopiContext(): WopiBootContext | null {
  if (typeof window === 'undefined') return null;
  const params = new URLSearchParams(window.location.search);
  const token = params.get('access_token');
  if (!token) return null;
  const payload = decodeJwtPayload(token);
  const fileId = typeof payload?.file_id === 'string' ? (payload.file_id as string) : null;
  if (!fileId) return null;
  return { accessToken: token, fileId };
}

function wopiUrl(path: string, token: string): string {
  // Path is "/wopi/files/<id>" or "/wopi/files/<id>/contents".
  // `access_token` lives in the query so the browser sends it on
  // both same-origin and (with the right CORS) cross-origin fetches
  // — same shape the WOPI spec defines.
  const sep = path.includes('?') ? '&' : '?';
  return `${path}${sep}access_token=${encodeURIComponent(token)}`;
}

const FETCH_NO_CREDS: RequestInit = {
  // WOPI auth is the URL token, not the personal-mode cookie. Omit
  // credentials so the two surfaces stay scope-isolated.
  credentials: 'omit',
};

export function createWopiFileSource(ctx: WopiBootContext): FileSource {
  let cachedRecent: RecentEntry | null = null;
  const subs = new Set<() => void>();
  const tick = () => {
    for (const fn of subs) {
      try {
        fn();
      } catch (err) {
        console.warn('[wopi-file-source] subscriber threw', err);
      }
    }
  };

  const fetchInfo = async (): Promise<CheckFileInfo> => {
    const res = await fetch(
      wopiUrl(`/wopi/files/${encodeURIComponent(ctx.fileId)}`, ctx.accessToken),
      {
        ...FETCH_NO_CREDS,
        method: 'GET',
      },
    );
    if (!res.ok) {
      throw new Error(`WOPI CheckFileInfo failed: HTTP ${res.status}`);
    }
    return (await res.json()) as CheckFileInfo;
  };

  return {
    kind: 'wopi',
    label: 'Embedded host',

    async listRecent(): Promise<RecentEntry[]> {
      if (cachedRecent) return [cachedRecent];
      try {
        const info = await fetchInfo();
        cachedRecent = {
          id: ctx.fileId,
          name: info.BaseFileName,
          sourceFormat: 'xlsx',
          size: info.Size,
          modifiedAt: Date.now(),
        };
        return [cachedRecent];
      } catch {
        // Network hiccup — return empty rather than throwing past
        // the home-screen render boundary.
        return [];
      }
    },

    subscribeRecent(cb) {
      subs.add(cb);
      return () => subs.delete(cb);
    },

    async openRecent(id): Promise<OpenedWorkbook> {
      if (id !== ctx.fileId) {
        throw new Error(`WOPI source bound to ${ctx.fileId}; refused to open ${id}`);
      }
      const res = await fetch(
        wopiUrl(`/wopi/files/${encodeURIComponent(ctx.fileId)}/contents`, ctx.accessToken),
        { ...FETCH_NO_CREDS, method: 'GET' },
      );
      if (!res.ok) {
        throw new Error(`WOPI GetFile failed: HTTP ${res.status}`);
      }
      const etag = res.headers.get('x-wopi-itemversion') ?? null;
      const buf = await res.arrayBuffer();
      const data = await xlsxToWorkbookData(buf);
      return {
        data,
        sourceFormat: 'xlsx',
        serverFileId: ctx.fileId,
        serverEtag: etag,
      };
    },

    async forgetRecent(): Promise<void> {
      // No-op. Embedded host owns its files; we can't remove on its
      // behalf. (A Disable / Hide command could land in a follow-up
      // if a host actually asks for it.)
    },

    async save(bytes: Blob, opts: SaveOptions): Promise<SaveResult> {
      // The WOPI token is bound to ctx.fileId — refuse to honour an
      // existingId that disagrees, rather than silently overwrite a
      // different file the URL token wouldn't authorise.
      const target = opts.existingId ?? ctx.fileId;
      if (target !== ctx.fileId) {
        throw new Error(`WOPI source bound to ${ctx.fileId}; refused to save to ${target}`);
      }
      const headers: Record<string, string> = {
        'content-type': 'application/octet-stream',
      };
      if (opts.existingEtag) headers['x-wopi-itemversion'] = opts.existingEtag;
      const buf = await bytes.arrayBuffer();
      const res = await fetch(
        wopiUrl(`/wopi/files/${encodeURIComponent(ctx.fileId)}/contents`, ctx.accessToken),
        {
          ...FETCH_NO_CREDS,
          method: 'POST',
          headers,
          body: buf,
        },
      );
      if (res.status === 409) {
        const body = (await res.json()) as { expected?: string };
        return { kind: 'conflict', expectedEtag: body.expected ?? '' };
      }
      if (!res.ok) {
        throw new Error(`WOPI PutFile failed: HTTP ${res.status}`);
      }
      const newEtag = res.headers.get('x-wopi-itemversion') ?? '';
      const info = await fetchInfo().catch(() => null);
      tick();
      return {
        kind: 'server',
        path: info?.BaseFileName ?? opts.filename,
        serverFileId: ctx.fileId,
        serverEtag: newEtag,
      };
    },
  };
}
