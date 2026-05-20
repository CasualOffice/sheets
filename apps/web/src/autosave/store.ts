import type { IWorkbookData } from '@univerjs/core';

/**
 * Tab-local autosave store. Single slot per origin — we keep the most
 * recent dirty snapshot so the user can recover from a tab-crash or an
 * accidental close. Anything more elaborate (versioned history,
 * cross-tab merge) is out of scope until we have a backend.
 *
 * IndexedDB is used over localStorage because workbooks regularly
 * exceed the 5 MB localStorage budget (one of our test fixtures is
 * already 4.8 MB pre-styling) and serializing/deserializing a JSON
 * string of that size every save is wasteful. IDB stores the object
 * directly via structured clone.
 *
 * The store deliberately writes to a single record (key 'current') so
 * the on-disk size stays bounded — old saves are overwritten. If we
 * ever want a list, switch to an autoincrement keyspace and add a
 * prune pass.
 */

const DB_NAME = 'casual-sheets';
const STORE = 'autosave';
const VERSION = 1;
const KEY = 'current';

export type AutosaveRecord = {
  /** Workbook name at the time of save — used in the restore prompt. */
  name: string;
  /** Source format (xlsx / ods / csv / tsv) if known, for round-trip on resave. */
  sourceFormat: string | null;
  /** Full snapshot the app can `replaceWorkbook` straight from. */
  data: IWorkbookData;
  /** Wall-clock ms at save time — for the restore prompt copy. */
  savedAt: number;
};

function openDb(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, VERSION);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(STORE)) db.createObjectStore(STORE);
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error ?? new Error('open db failed'));
  });
}

export async function readAutosave(): Promise<AutosaveRecord | null> {
  try {
    const db = await openDb();
    return await new Promise<AutosaveRecord | null>((resolve, reject) => {
      const tx = db.transaction(STORE, 'readonly');
      const req = tx.objectStore(STORE).get(KEY);
      req.onsuccess = () => resolve((req.result as AutosaveRecord | undefined) ?? null);
      req.onerror = () => reject(req.error ?? new Error('read failed'));
    });
  } catch (err) {
    console.warn('[autosave] read failed; skipping restore', err);
    return null;
  }
}

export async function writeAutosave(rec: AutosaveRecord): Promise<void> {
  const db = await openDb();
  await new Promise<void>((resolve, reject) => {
    const tx = db.transaction(STORE, 'readwrite');
    tx.objectStore(STORE).put(rec, KEY);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error ?? new Error('write failed'));
  });
}

export async function clearAutosave(): Promise<void> {
  try {
    const db = await openDb();
    await new Promise<void>((resolve, reject) => {
      const tx = db.transaction(STORE, 'readwrite');
      tx.objectStore(STORE).delete(KEY);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error ?? new Error('clear failed'));
    });
  } catch (err) {
    console.warn('[autosave] clear failed', err);
  }
}
