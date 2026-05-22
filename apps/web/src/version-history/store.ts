import type { IWorkbookData } from '@univerjs/core';

/**
 * Local version-history snapshot store. Separate object store from the
 * single-slot autosave (`apps/web/src/autosave/store.ts`) — the two
 * solve different problems:
 *
 *   - **Autosave**: one slot, overwritten on every change, used for
 *     "I crashed mid-edit — recover the workbook." Lifetime: until the
 *     user explicitly saves or discards.
 *   - **Version history**: many slots, captured at coarse-grained
 *     moments (every ~10 min while dirty, or on explicit "Save
 *     version"), used for "I want to roll back to how it looked an
 *     hour ago." Lifetime: kept until pruned or hand-deleted.
 *
 * Retention rules:
 *   - Manual snapshots are NEVER auto-pruned — the user asked for them
 *     to exist. Only an explicit delete removes one.
 *   - Auto snapshots are kept by count: `AUTO_RETENTION` newest auto
 *     entries survive each write. Earlier ones are dropped.
 *
 * Storage cost: each snapshot is a full IWorkbookData. A typical
 * 1k-cell workbook serialises to ~50 KB; a 100k-cell sheet can be a
 * few MB. With 30 auto + N manual snapshots a power user can pass
 * 100 MB of IDB use. Future work: lz-string compression on write to
 * cut that ~4×.
 */

import type { LiveVersionFeed } from './live-feed';

const DB_NAME = 'casual-sheets';
const STORE = 'versions';
// Bumped from autosave's `1` because we add a new object store. The
// upgrade handler is additive — existing `autosave` data is preserved.
const VERSION = 2;
const AUTO_RETENTION = 30;

export type VersionKind = 'auto' | 'manual';

export type VersionSnapshot = {
  /** IDB auto-increment key, assigned on first persist. Optional for
   *  the in-memory `VersionDraft` shape the capture hook hands in. */
  id?: number;
  kind: VersionKind;
  /** User-supplied label for manual versions; derived (e.g. "Auto-save")
   *  for auto versions — kept on the record so list rendering doesn't
   *  have to branch. */
  name: string;
  /** Wall-clock ms at capture. Used for time grouping in the panel. */
  savedAt: number;
  /** Source format at capture, for round-trip on later "Save as" from a
   *  preview. Mirrors the autosave shape. */
  sourceFormat: string | null;
  /** Full workbook snapshot — replaceWorkbook ready. */
  data: IWorkbookData;
  /** Approximate JSON byte size, set on persist for UI display. */
  size?: number;
};

export type VersionDraft = Omit<VersionSnapshot, 'id'>;

function openDb(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, VERSION);
    req.onupgradeneeded = (event) => {
      const db = req.result;
      // From v0 (no DB yet): create the autosave store as well, so a
      // fresh install gets both. From v1 (autosave already there):
      // only the versions store is new.
      if (!db.objectStoreNames.contains('autosave')) {
        db.createObjectStore('autosave');
      }
      if (!db.objectStoreNames.contains(STORE)) {
        const os = db.createObjectStore(STORE, { keyPath: 'id', autoIncrement: true });
        // Index on savedAt for "give me newest N" queries used by the
        // panel's listing without scanning.
        os.createIndex('savedAt', 'savedAt', { unique: false });
        // Index on kind so retention pruning can iterate auto entries
        // without re-reading the whole store.
        os.createIndex('kind', 'kind', { unique: false });
      }
      // Touch event to keep TS happy about unused param.
      void event;
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error ?? new Error('open db failed'));
  });
}

/**
 * Persist a draft snapshot. Returns the assigned id. Triggers a prune
 * pass that drops the oldest `auto` entries past `AUTO_RETENTION`.
 */
export async function writeVersion(draft: VersionDraft): Promise<number> {
  const db = await openDb();
  const size = estimateSize(draft.data);
  const record: VersionDraft = { ...draft, size };
  const id = await new Promise<number>((resolve, reject) => {
    const tx = db.transaction(STORE, 'readwrite');
    const req = tx.objectStore(STORE).add(record);
    req.onsuccess = () => resolve(req.result as number);
    req.onerror = () => reject(req.error ?? new Error('write failed'));
  });
  // Fire-and-forget retention sweep; failures are non-fatal — old
  // entries are cosmetic clutter, not correctness.
  void pruneAuto().catch((err) => console.warn('[version-history] prune failed', err));
  notifyFeed();
  return id;
}

export async function listVersions(): Promise<VersionSnapshot[]> {
  try {
    const db = await openDb();
    return await new Promise<VersionSnapshot[]>((resolve, reject) => {
      const tx = db.transaction(STORE, 'readonly');
      const req = tx.objectStore(STORE).getAll();
      req.onsuccess = () => {
        const list = (req.result ?? []) as VersionSnapshot[];
        // Newest first.
        list.sort((a, b) => b.savedAt - a.savedAt);
        resolve(list);
      };
      req.onerror = () => reject(req.error ?? new Error('list failed'));
    });
  } catch (err) {
    console.warn('[version-history] list failed', err);
    return [];
  }
}

export async function readVersion(id: number): Promise<VersionSnapshot | null> {
  try {
    const db = await openDb();
    return await new Promise<VersionSnapshot | null>((resolve, reject) => {
      const tx = db.transaction(STORE, 'readonly');
      const req = tx.objectStore(STORE).get(id);
      req.onsuccess = () => resolve((req.result as VersionSnapshot | undefined) ?? null);
      req.onerror = () => reject(req.error ?? new Error('read failed'));
    });
  } catch (err) {
    console.warn('[version-history] read failed', err);
    return null;
  }
}

export async function renameVersion(id: number, name: string): Promise<void> {
  const db = await openDb();
  await new Promise<void>((resolve, reject) => {
    const tx = db.transaction(STORE, 'readwrite');
    const os = tx.objectStore(STORE);
    const get = os.get(id);
    get.onsuccess = () => {
      const existing = get.result as VersionSnapshot | undefined;
      if (!existing) {
        resolve();
        return;
      }
      os.put({ ...existing, name });
    };
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error ?? new Error('rename failed'));
  });
  notifyFeed();
}

export async function deleteVersion(id: number): Promise<void> {
  const db = await openDb();
  await new Promise<void>((resolve, reject) => {
    const tx = db.transaction(STORE, 'readwrite');
    tx.objectStore(STORE).delete(id);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error ?? new Error('delete failed'));
  });
  notifyFeed();
}

/** Drop everything. Used by tests + the user's "Clear version history"
 *  action in the panel (yet-to-be-added). */
export async function clearAllVersions(): Promise<void> {
  const db = await openDb();
  await new Promise<void>((resolve, reject) => {
    const tx = db.transaction(STORE, 'readwrite');
    tx.objectStore(STORE).clear();
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error ?? new Error('clear failed'));
  });
  notifyFeed();
}

async function pruneAuto(): Promise<void> {
  const db = await openDb();
  // Walk the `kind` index for 'auto' entries, sorted by their built-in
  // primary key (insertion order — also ascending by id), and delete
  // anything past the retention window.
  await new Promise<void>((resolve, reject) => {
    const tx = db.transaction(STORE, 'readwrite');
    const os = tx.objectStore(STORE);
    const idx = os.index('kind');
    const range = IDBKeyRange.only('auto');
    const req = idx.openCursor(range, 'prev');
    let kept = 0;
    req.onsuccess = () => {
      const cursor = req.result;
      if (!cursor) {
        resolve();
        return;
      }
      kept += 1;
      if (kept > AUTO_RETENTION) {
        os.delete(cursor.primaryKey);
      }
      cursor.continue();
    };
    req.onerror = () => reject(req.error ?? new Error('prune failed'));
    tx.oncomplete = () => resolve();
  });
}

/** Crude byte-size estimate — JSON-stringify length in chars, treated
 *  as a near-1:1 byte proxy. Fine for the UI's "≈ 120 KB" hint. */
function estimateSize(data: IWorkbookData): number {
  try {
    return JSON.stringify(data).length;
  } catch {
    return 0;
  }
}

let feed: LiveVersionFeed | null = null;
export function setLiveFeed(f: LiveVersionFeed | null) {
  feed = f;
}
function notifyFeed() {
  feed?.tick();
}
