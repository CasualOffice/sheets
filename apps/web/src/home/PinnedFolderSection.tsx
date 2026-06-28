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

import { useEffect, useState } from 'react';
import { Icon } from '../shell/Icon';
import {
  listSpreadsheetEntries,
  reconnectFolder,
  type FolderState,
} from '../file-system-access/pinned-folder';

/**
 * "From your folder" section on the home screen — lists the
 * spreadsheet entries directly under the pinned folder, click to
 * open. Distinct from "Recent files" (IDB-backed history) — these
 * are real files on disk that the user has explicitly chosen to
 * surface here.
 *
 * Only renders when the pinned folder is `granted` or `prompt`; when
 * `prompt` we show a one-tap "Reconnect" affordance instead of a stale
 * empty list. Caller passes the state + a callback so the workbook-
 * swap stays in the host component.
 */

type Entry = { name: string; size: number; handle: FileSystemFileHandle };

export function PinnedFolderSection({
  state,
  onReconnect,
  onOpenFile,
}: {
  state: Extract<FolderState, { kind: 'granted' | 'prompt' }>;
  onReconnect: () => void;
  onOpenFile: (file: File) => Promise<void>;
}) {
  const [entries, setEntries] = useState<Entry[] | null>(null);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    if (state.kind !== 'granted') {
      setEntries(null);
      return;
    }
    let cancelled = false;
    void (async () => {
      try {
        const raw = await listSpreadsheetEntries(state.record.handle);
        const sized: Entry[] = [];
        // Read file sizes lazily so a 50-entry folder doesn't open 50
        // files. `getFile()` is cheap (metadata only), but iterating in
        // sequence keeps the UI responsive.
        for (const r of raw) {
          try {
            const file = await r.handle.getFile();
            sized.push({ name: r.name, size: file.size, handle: r.handle });
          } catch {
            sized.push({ name: r.name, size: 0, handle: r.handle });
          }
        }
        if (!cancelled) setEntries(sized);
      } catch (err) {
        if (!cancelled) {
          setError(err instanceof Error ? err.message : 'Could not list folder');
        }
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [state]);

  return (
    <section className="home__section home__section--pinned" data-testid="home-pinned-section">
      <div className="home__section-head">
        <h2>
          <Icon name="folder_special" /> From {state.record.name}
        </h2>
        <span className="home__section-hint">
          {state.kind === 'granted'
            ? 'Files on your computer — opened in-browser.'
            : 'Reconnect to list files.'}
        </span>
      </div>

      {state.kind === 'prompt' && (
        <div className="home__pinned-empty">
          <span>Folder access expired with the last browser session.</span>
          <button
            type="button"
            className="home__reopen-btn home__reopen-btn--primary"
            onClick={async () => {
              const ok = await reconnectFolder(state.record);
              if (ok) onReconnect();
            }}
            data-testid="home-pinned-reconnect"
          >
            Reconnect
          </button>
        </div>
      )}

      {state.kind === 'granted' && entries === null && !error && (
        <div className="home__pinned-empty">Reading folder…</div>
      )}

      {state.kind === 'granted' && error && (
        <div className="home__pinned-empty">Couldn’t read folder: {error}</div>
      )}

      {state.kind === 'granted' && entries && entries.length === 0 && (
        <div className="home__pinned-empty">
          No .xlsx, .ods, .csv, or .tsv files in this folder yet.
        </div>
      )}

      {state.kind === 'granted' && entries && entries.length > 0 && (
        <ul className="home__recents">
          {entries.map((e) => (
            <li key={e.name} className="home__recent">
              <button
                type="button"
                className="home__recent-open"
                data-testid="home-pinned-open"
                onClick={async () => {
                  const file = await e.handle.getFile();
                  await onOpenFile(file);
                }}
              >
                <span className="home__recent-icon">
                  <Icon name="description" />
                </span>
                <span className="home__recent-text">
                  <span className="home__recent-name">{e.name}</span>
                  <span className="home__recent-meta">{formatSize(e.size)}</span>
                </span>
              </button>
            </li>
          ))}
        </ul>
      )}
    </section>
  );
}

function formatSize(bytes: number): string {
  if (!bytes) return '';
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(0)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}
