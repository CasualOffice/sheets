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

/**
 * IDB-backed persistence for the pinned `FileSystemDirectoryHandle`.
 *
 * IndexedDB is the only browser storage that survives reloads AND can
 * hold a directory handle — `localStorage` only takes strings and would
 * lose the live OS-level reference. The handle itself is stored under a
 * single fixed key so re-pinning replaces the previous value.
 *
 * Permission lifecycle: the handle survives indefinitely but the
 * granted permission on it is dropped at the end of the browsing
 * session. On reload we need to call `requestPermission()` again before
 * any read/write succeeds. See `pinned-folder.ts` for the flow.
 */

const DB_NAME = 'casual-sheets';
const STORE = 'pinned-folder';
const KEY = 'current';
// Shared DB version — must equal autosave/store.ts, version-history/store.ts,
// and recent-files/store.ts. Bumped from 3 to add this store.
const VERSION = 4;

function openDb(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, VERSION);
    req.onupgradeneeded = () => {
      const db = req.result;
      // Idempotent — provisions every shared store so a fresh install
      // ends up with a complete schema regardless of which module
      // opens the DB first.
      if (!db.objectStoreNames.contains('autosave')) {
        db.createObjectStore('autosave');
      }
      if (!db.objectStoreNames.contains('versions')) {
        const os = db.createObjectStore('versions', { keyPath: 'id', autoIncrement: true });
        os.createIndex('savedAt', 'savedAt', { unique: false });
        os.createIndex('kind', 'kind', { unique: false });
      }
      if (!db.objectStoreNames.contains('recent-files')) {
        const os = db.createObjectStore('recent-files', { keyPath: 'id', autoIncrement: true });
        os.createIndex('openedAt', 'openedAt', { unique: false });
        os.createIndex('name', 'name', { unique: false });
      }
      if (!db.objectStoreNames.contains(STORE)) {
        db.createObjectStore(STORE);
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error ?? new Error('open db failed'));
  });
}

export type PinnedRecord = {
  handle: FileSystemDirectoryHandle;
  /** Display name from `handle.name`, copied here so the UI can render
   *  it without waiting for a permission grant on the handle itself. */
  name: string;
  pinnedAt: number;
};

export async function readPinnedFolder(): Promise<PinnedRecord | null> {
  try {
    const db = await openDb();
    return await new Promise<PinnedRecord | null>((resolve, reject) => {
      const tx = db.transaction(STORE, 'readonly');
      const req = tx.objectStore(STORE).get(KEY);
      req.onsuccess = () => resolve((req.result as PinnedRecord | undefined) ?? null);
      req.onerror = () => reject(req.error ?? new Error('read failed'));
    });
  } catch (err) {
    console.warn('[fsa] read pinned folder failed', err);
    return null;
  }
}

export async function writePinnedFolder(rec: PinnedRecord): Promise<void> {
  const db = await openDb();
  await new Promise<void>((resolve, reject) => {
    const tx = db.transaction(STORE, 'readwrite');
    tx.objectStore(STORE).put(rec, KEY);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error ?? new Error('write failed'));
  });
}

export async function clearPinnedFolder(): Promise<void> {
  const db = await openDb();
  await new Promise<void>((resolve, reject) => {
    const tx = db.transaction(STORE, 'readwrite');
    tx.objectStore(STORE).delete(KEY);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error ?? new Error('delete failed'));
  });
}
