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

import type { IWorkbookData } from '@univerjs/core';
import {
  deleteRecentFile,
  listRecentFiles,
  setLiveFeed,
  type RecentFile,
} from '../recent-files/store';
import { createLiveRecentFilesFeed } from '../recent-files/live-feed';
import { getFolderState, writeFileToFolder } from '../file-system-access/pinned-folder';
import type { FileSource, OpenedWorkbook, RecentEntry, SaveOptions, SaveResult } from './types';

/**
 * Phase B `FileSource` for the GitHub-Pages / browser-only deploy.
 *
 * Composes the existing modules — recent-files IDB store, the
 * `LiveRecentFilesFeed`, the pinned-folder API, the download-blob
 * trigger — behind the unified `FileSource` shape. Nothing here
 * changes user-visible behaviour; it just relocates the dispatch so
 * Mode 3 and Mode 2 can later swap in their own source against the
 * same contract.
 *
 * Why a class vs a plain object: the live-feed is lazily initialised
 * (creating it imports the IDB module). A factory + cached singleton
 * keeps boot light without giving up identity-based equality for
 * React context.
 */

function recentFromRecord(rec: RecentFile): RecentEntry {
  return {
    // IDB autoincrement → string for the opaque-id contract.
    id: String(rec.id ?? ''),
    name: rec.name,
    sourceFormat: rec.sourceFormat,
    size: rec.size,
    modifiedAt: rec.openedAt,
  };
}

/** Browser-only delivery: write directly to a pinned FSA folder when
 *  granted, otherwise trigger a download. Equivalent to the previous
 *  `deliverBlob` in `file-actions` — the call site moved here so
 *  swapping the source swaps the delivery channel automatically. */
async function deliverBrowserBlob(blob: Blob, filename: string): Promise<SaveResult> {
  const state = await getFolderState();
  if (state.kind === 'granted') {
    await writeFileToFolder(state.record.handle, filename, blob);
    return { kind: 'folder', folderName: state.record.name };
  }
  triggerDownload(blob, filename);
  return { kind: 'download' };
}

function triggerDownload(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 0);
}

export function createBrowserFileSource(): FileSource {
  // Lazily provision the live-feed so importing this module doesn't
  // touch the recent-files store until something actually subscribes.
  let feed: ReturnType<typeof createLiveRecentFilesFeed> | null = null;
  const ensureFeed = () => {
    if (!feed) {
      feed = createLiveRecentFilesFeed();
      setLiveFeed(feed);
    }
    return feed;
  };

  return {
    kind: 'browser',
    label: 'This browser',

    async listRecent(): Promise<RecentEntry[]> {
      const raw = await listRecentFiles();
      return raw.map(recentFromRecord);
    },

    subscribeRecent(cb) {
      return ensureFeed().subscribe(cb);
    },

    async openRecent(id): Promise<OpenedWorkbook> {
      const raw = await listRecentFiles();
      const match = raw.find((r) => String(r.id ?? '') === id);
      if (!match) throw new Error(`recent file not found: ${id}`);
      return {
        data: match.data as IWorkbookData,
        sourceFormat: match.sourceFormat,
      };
    },

    async forgetRecent(id) {
      const numeric = Number(id);
      if (!Number.isFinite(numeric)) return;
      await deleteRecentFile(numeric);
    },

    save(bytes: Blob, opts: SaveOptions) {
      return deliverBrowserBlob(bytes, opts.filename);
    },
  };
}
