import { isFsaSupported } from './support';
import { clearPinnedFolder, readPinnedFolder, writePinnedFolder, type PinnedRecord } from './store';

/**
 * High-level API for the "pinned folder" feature — the user picks a
 * folder once, and subsequent saves write directly into it instead of
 * triggering a browser download.
 *
 * Three permission states we care about, returned by `verifyPermission`:
 *
 *   - 'granted'    — handle exists + readwrite permission is live; safe
 *                    to read/write without prompting.
 *   - 'prompt'     — handle exists but permission expired (browsing
 *                    session ended); caller must request a re-grant on
 *                    a user gesture.
 *   - 'denied'     — handle exists but the user has actively denied. We
 *                    treat this as 'not pinned' from the UI's
 *                    perspective and offer to re-pin.
 *   - 'unsupported'— browser without the API (Firefox, Safari). UI
 *                    hides the pin controls entirely.
 *   - 'none'       — no folder has been pinned.
 */

export type FolderState =
  | { kind: 'unsupported' }
  | { kind: 'none' }
  | { kind: 'granted'; record: PinnedRecord }
  | { kind: 'prompt'; record: PinnedRecord }
  | { kind: 'denied'; record: PinnedRecord };

export async function getFolderState(): Promise<FolderState> {
  if (!isFsaSupported()) return { kind: 'unsupported' };
  const record = await readPinnedFolder();
  if (!record) return { kind: 'none' };
  const status = await queryPermission(record.handle);
  if (status === 'granted') return { kind: 'granted', record };
  if (status === 'denied') return { kind: 'denied', record };
  return { kind: 'prompt', record };
}

/**
 * Walk the user through `showDirectoryPicker` and persist the chosen
 * handle. Must be called from a user gesture (click) — browsers gate
 * the picker behind a user-activation check.
 *
 * Returns the new pinned record, or `null` if the user cancelled.
 */
export async function pinFolder(): Promise<PinnedRecord | null> {
  if (!isFsaSupported()) return null;
  try {
    const handle = await (
      window as unknown as {
        showDirectoryPicker: (opts?: {
          mode?: 'read' | 'readwrite';
        }) => Promise<FileSystemDirectoryHandle>;
      }
    ).showDirectoryPicker({ mode: 'readwrite' });
    const record: PinnedRecord = {
      handle,
      name: handle.name,
      pinnedAt: Date.now(),
    };
    await writePinnedFolder(record);
    return record;
  } catch (err) {
    // AbortError fires when the user dismisses the picker. Anything
    // else (SecurityError, NotAllowedError) we surface so the caller
    // can show a toast.
    if (err instanceof DOMException && err.name === 'AbortError') return null;
    throw err;
  }
}

/**
 * Re-request `readwrite` permission on the stored handle. Must be
 * called from a user gesture. Returns true on grant.
 */
export async function reconnectFolder(record: PinnedRecord): Promise<boolean> {
  if (!isFsaSupported()) return false;
  const result = await requestPermission(record.handle);
  return result === 'granted';
}

export async function unpinFolder(): Promise<void> {
  await clearPinnedFolder();
}

/**
 * Write a workbook blob to the pinned folder under `filename`. Replaces
 * any existing file with the same name. Caller is responsible for the
 * filename — including the extension.
 *
 * Returns the FileSystemFileHandle so the caller can re-save to the
 * same file without going through the directory again (saves a
 * permission check + a name resolution per save in a hot loop).
 */
export async function writeFileToFolder(
  handle: FileSystemDirectoryHandle,
  filename: string,
  data: Blob,
): Promise<FileSystemFileHandle> {
  const fileHandle = await handle.getFileHandle(filename, { create: true });
  const writable = await fileHandle.createWritable();
  try {
    await writable.write(data);
  } finally {
    await writable.close();
  }
  return fileHandle;
}

/** List `.xlsx`/`.ods`/`.csv`/`.tsv` entries directly under the folder
 *  (non-recursive — flat workbook folder is the model we promote). */
export async function listSpreadsheetEntries(
  handle: FileSystemDirectoryHandle,
): Promise<Array<{ name: string; handle: FileSystemFileHandle }>> {
  const out: Array<{ name: string; handle: FileSystemFileHandle }> = [];
  const dirHandle = handle as FileSystemDirectoryHandle & {
    values: () => AsyncIterableIterator<FileSystemHandle>;
  };
  for await (const entry of dirHandle.values()) {
    if (entry.kind !== 'file') continue;
    if (!/\.(xlsx|xlsm|ods|csv|tsv|tab)$/i.test(entry.name)) continue;
    out.push({ name: entry.name, handle: entry as FileSystemFileHandle });
  }
  out.sort((a, b) => a.name.localeCompare(b.name));
  return out;
}

type PermissionStatus = 'granted' | 'denied' | 'prompt';

async function queryPermission(handle: FileSystemDirectoryHandle): Promise<PermissionStatus> {
  const h = handle as FileSystemDirectoryHandle & {
    queryPermission?: (opts: { mode: 'read' | 'readwrite' }) => Promise<PermissionStatus>;
  };
  if (typeof h.queryPermission !== 'function') return 'prompt';
  try {
    return await h.queryPermission({ mode: 'readwrite' });
  } catch {
    return 'prompt';
  }
}

async function requestPermission(handle: FileSystemDirectoryHandle): Promise<PermissionStatus> {
  const h = handle as FileSystemDirectoryHandle & {
    requestPermission?: (opts: { mode: 'read' | 'readwrite' }) => Promise<PermissionStatus>;
  };
  if (typeof h.requestPermission !== 'function') return 'prompt';
  try {
    return await h.requestPermission({ mode: 'readwrite' });
  } catch {
    return 'denied';
  }
}
