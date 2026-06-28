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
 * Server-backed `FileSource` for the personal-mode docker (Mode 3
 * of #49). Talks to the `/files/*` routes added in Phase C Batch 2;
 * the byte payload still lives in the configured HostIntegration
 * backend (local / s3 / postgres), the session cookie tags along.
 *
 * Phase C minimum viable shape:
 *
 *   - listRecent      — GET /files
 *   - openRecent      — GET /files/:id → parse via xlsxToWorkbookData
 *   - forgetRecent    — DELETE /files/:id
 *   - save            — POST /files (multipart upload as a new
 *                       entry). In-place "Save" against an existing
 *                       id + If-Match etag for the conflict modal
 *                       lands in a follow-up — for v1 every Save is
 *                       a new entry in My Files, so users never lose
 *                       data to a race.
 *   - subscribeRecent — `ticker` callback registry the file-actions
 *                       layer pokes after a successful save, so the
 *                       home-screen recents strip refreshes.
 *
 * Network errors surface as plain `Error`s and propagate to the
 * caller (`MenuBar.exportAs` wraps Save in a toast.error). 401
 * responses are treated as "session expired" — caller resets the
 * AuthState so the gate re-renders the login view.
 */

type Ticker = {
  tick: () => void;
  subscribe: (cb: () => void) => () => void;
};

function createTicker(): Ticker {
  const subs = new Set<() => void>();
  return {
    tick: () => {
      for (const fn of subs) {
        try {
          fn();
        } catch (err) {
          console.warn('[personal-file-source] subscriber threw', err);
        }
      }
    },
    subscribe: (fn) => {
      subs.add(fn);
      return () => subs.delete(fn);
    },
  };
}

type ServerFileMeta = {
  id: string;
  name: string;
  size: number;
  etag: string;
  createdAt: number;
  modifiedAt: number;
};

const COMMON: RequestInit = { credentials: 'include' };

export function createPersonalFileSource(label = 'My files'): FileSource {
  const ticker = createTicker();

  const guardOk = async (res: Response, op: string): Promise<void> => {
    if (res.ok) return;
    // 401: caller should reset auth (gate re-renders). We throw a
    // typed Error so the caller can `instanceof PersonalAuthExpired`
    // if it wants to, without coupling to status codes.
    if (res.status === 401) throw new PersonalAuthExpired();
    let body: { error?: string } | null = null;
    try {
      body = (await res.json()) as { error?: string };
    } catch {
      /* not json */
    }
    throw new Error(`${op}: HTTP ${res.status}${body?.error ? ` (${body.error})` : ''}`);
  };

  return {
    kind: 'personal',
    label,

    async listRecent(): Promise<RecentEntry[]> {
      const res = await fetch('/files', { ...COMMON });
      await guardOk(res, 'list files');
      const body = (await res.json()) as { files: ServerFileMeta[] };
      return body.files.map((f) => ({
        id: f.id,
        name: f.name,
        // Server side has no `sourceFormat` concept — every personal
        // file goes in / comes out as xlsx (the format we control).
        sourceFormat: 'xlsx',
        size: f.size,
        modifiedAt: f.modifiedAt,
      }));
    },

    subscribeRecent(cb) {
      return ticker.subscribe(cb);
    },

    async openRecent(id): Promise<OpenedWorkbook> {
      const res = await fetch(`/files/${encodeURIComponent(id)}`, { ...COMMON });
      await guardOk(res, 'open file');
      const etag = res.headers.get('etag') ?? null;
      const buf = await res.arrayBuffer();
      const data = await xlsxToWorkbookData(buf);
      return { data, sourceFormat: 'xlsx', serverFileId: id, serverEtag: etag };
    },

    async forgetRecent(id): Promise<void> {
      const res = await fetch(`/files/${encodeURIComponent(id)}`, {
        ...COMMON,
        method: 'DELETE',
      });
      await guardOk(res, 'delete file');
      ticker.tick();
    },

    async save(bytes: Blob, opts: SaveOptions): Promise<SaveResult> {
      // In-place save when the caller tracked the file id from
      // openRecent: POST /files/:id with raw bytes + If-Match header.
      // Server returns 412 on stale etag — surface as a conflict so
      // file-actions can pop the shared modal.
      if (opts.existingId) {
        const headers: Record<string, string> = {
          'content-type': 'application/octet-stream',
        };
        if (opts.existingEtag) headers['if-match'] = opts.existingEtag;
        const buf = await bytes.arrayBuffer();
        const res = await fetch(`/files/${encodeURIComponent(opts.existingId)}`, {
          ...COMMON,
          method: 'POST',
          headers,
          body: buf,
        });
        if (res.status === 412) {
          const body = (await res.json()) as { expected?: string };
          return { kind: 'conflict', expectedEtag: body.expected ?? '' };
        }
        await guardOk(res, 'save file (in place)');
        const body = (await res.json()) as { file: ServerFileMeta };
        ticker.tick();
        return {
          kind: 'server',
          path: body.file.name,
          serverFileId: body.file.id,
          serverEtag: body.file.etag,
        };
      }
      // Fall-through: brand-new save (no existingId tracked).
      // Multipart POST as before — creates a new entry in My Files.
      const form = new FormData();
      form.append('file', bytes, opts.filename);
      form.append('name', opts.filename);
      const res = await fetch('/files', { ...COMMON, method: 'POST', body: form });
      await guardOk(res, 'save file');
      const body = (await res.json()) as { file: ServerFileMeta };
      ticker.tick();
      return {
        kind: 'server',
        path: body.file.name,
        serverFileId: body.file.id,
        serverEtag: body.file.etag,
      };
    },
  };
}

/** Thrown by any `/files/*` call when the server returns 401. The
 *  caller (typically the AuthProvider) refreshes the state so the
 *  gate re-renders the login view. */
export class PersonalAuthExpired extends Error {
  constructor() {
    super('session expired');
    this.name = 'PersonalAuthExpired';
  }
}
