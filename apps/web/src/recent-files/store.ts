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
import type { WorkbookFormat } from '../workbook-context';

/**
 * IndexedDB-backed "recent files" list. Captures the user's last N
 * opened workbooks so the landing screen can let them reopen with
 * one click. Distinct from autosave (single recovery slot) and
 * version history (per-workbook timeline) — this is "what files have
 * I been working with lately, across sessions".
 *
 * Retention: capped at MAX_ENTRIES, oldest evicted first. We also
 * drop entries that haven't been opened in STALE_AFTER_MS to keep
 * the listing fresh.
 */

const DB_NAME = 'casual-sheets';
const STORE = 'recent-files';
// Shared DB version — must equal autosave, version-history,
// file-system-access stores. Bumped from 3 to add `pinned-folder`.
const VERSION = 4;
const MAX_ENTRIES = 10;
const STALE_AFTER_MS = 60 * 24 * 60 * 60 * 1000; // 60 days

export type RecentFile = {
  /** IDB auto-increment key. Optional at insert time. */
  id?: number;
  name: string;
  sourceFormat: WorkbookFormat | null;
  data: IWorkbookData;
  size: number;
  openedAt: number;
};

import type { LiveRecentFilesFeed } from './live-feed';

function openDb(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, VERSION);
    req.onupgradeneeded = () => {
      const db = req.result;
      // v1: autosave store. v2: versions store. v3: recent-files store.
      // The handler is additive — each existing store is preserved.
      if (!db.objectStoreNames.contains('autosave')) {
        db.createObjectStore('autosave');
      }
      if (!db.objectStoreNames.contains('versions')) {
        const os = db.createObjectStore('versions', { keyPath: 'id', autoIncrement: true });
        os.createIndex('savedAt', 'savedAt', { unique: false });
        os.createIndex('kind', 'kind', { unique: false });
      }
      if (!db.objectStoreNames.contains(STORE)) {
        const os = db.createObjectStore(STORE, { keyPath: 'id', autoIncrement: true });
        os.createIndex('openedAt', 'openedAt', { unique: false });
        os.createIndex('name', 'name', { unique: false });
      }
      if (!db.objectStoreNames.contains('pinned-folder')) {
        db.createObjectStore('pinned-folder');
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error ?? new Error('open db failed'));
  });
}

let feed: LiveRecentFilesFeed | null = null;
export function setLiveFeed(f: LiveRecentFilesFeed | null) {
  feed = f;
}
function notifyFeed() {
  feed?.tick();
}

/**
 * Record (or refresh) a workbook in the recent-files list. If an entry
 * with the same name already exists, it's updated in place; otherwise
 * a new entry is appended. Triggers a prune pass to enforce the cap.
 */
export async function recordRecentFile(rec: Omit<RecentFile, 'id'>): Promise<void> {
  const db = await openDb();
  const size = rec.size || estimateSize(rec.data);
  await new Promise<void>((resolve, reject) => {
    const tx = db.transaction(STORE, 'readwrite');
    const os = tx.objectStore(STORE);
    const idx = os.index('name');
    const getReq = idx.get(rec.name);
    getReq.onsuccess = () => {
      const existing = getReq.result as RecentFile | undefined;
      if (existing && existing.id != null) {
        os.put({ ...existing, ...rec, id: existing.id, size });
      } else {
        os.add({ ...rec, size });
      }
    };
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error ?? new Error('record failed'));
  });
  void prune().catch((err) => console.warn('[recent-files] prune failed', err));
  notifyFeed();
}

export async function listRecentFiles(): Promise<RecentFile[]> {
  try {
    const db = await openDb();
    const all = await new Promise<RecentFile[]>((resolve, reject) => {
      const tx = db.transaction(STORE, 'readonly');
      const req = tx.objectStore(STORE).getAll();
      req.onsuccess = () => resolve((req.result ?? []) as RecentFile[]);
      req.onerror = () => reject(req.error ?? new Error('list failed'));
    });
    const cutoff = Date.now() - STALE_AFTER_MS;
    return all.filter((r) => r.openedAt >= cutoff).sort((a, b) => b.openedAt - a.openedAt);
  } catch (err) {
    console.warn('[recent-files] list failed', err);
    return [];
  }
}

export async function deleteRecentFile(id: number): Promise<void> {
  const db = await openDb();
  await new Promise<void>((resolve, reject) => {
    const tx = db.transaction(STORE, 'readwrite');
    tx.objectStore(STORE).delete(id);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error ?? new Error('delete failed'));
  });
  notifyFeed();
}

async function prune(): Promise<void> {
  const db = await openDb();
  await new Promise<void>((resolve, reject) => {
    const tx = db.transaction(STORE, 'readwrite');
    const os = tx.objectStore(STORE);
    const idx = os.index('openedAt');
    const req = idx.openCursor(null, 'prev');
    let kept = 0;
    req.onsuccess = () => {
      const cursor = req.result;
      if (!cursor) return;
      kept += 1;
      // Drop everything past MAX_ENTRIES, plus anything stale.
      if (
        kept > MAX_ENTRIES ||
        (cursor.value as RecentFile).openedAt < Date.now() - STALE_AFTER_MS
      ) {
        os.delete(cursor.primaryKey);
      }
      cursor.continue();
    };
    req.onerror = () => reject(req.error ?? new Error('prune failed'));
    tx.oncomplete = () => resolve();
  });
}

function estimateSize(data: IWorkbookData): number {
  try {
    return JSON.stringify(data).length;
  } catch {
    return 0;
  }
}
