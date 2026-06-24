import type { IWorkbookData } from '@univerjs/core';
import type { WorkbookFormat } from '../workbook-context';

/**
 * The web app's storage-IO contract.
 *
 * One interface, three implementations (today: Browser only; Mode 3 and
 * Mode 2 land in Phases C and D). Every part of the UI that needs to
 * list, open, or save a workbook talks to a `FileSource` instead of
 * reaching into IDB, the FSA module, or a server route directly. That
 * way Mode 3 (Personal account on a docker volume) and Mode 2 (WOPI
 * embed) can swap in their own implementation without the File menu,
 * landing screen, or save flow learning anything new.
 *
 * Design notes
 *
 *   - **Identifiers are opaque strings.** Browser uses the IDB
 *     auto-increment key serialised to a string; server modes use
 *     whatever id the host integration assigns. Callers never parse
 *     them.
 *   - **Workbook bytes vs snapshots.** `openRecent` returns a full
 *     `IWorkbookData` because Mode 1's recent-files store keeps the
 *     parsed snapshot in IDB (avoids a re-parse on reopen). Mode 3 /
 *     Mode 2 implementations will fetch bytes, parse them, and return
 *     the same shape — the parse stays inside the source so callers
 *     don't have to branch.
 *   - **`save` takes a Blob.** The worker-level xlsx / ods / csv / tsv
 *     serialization stays in `file-actions`; the source's job is just
 *     to deliver the bytes (download / FSA / WOPI PUT).
 *   - **`SaveResult` is discriminated** so the caller can adapt the
 *     toast copy without coupling to the source.
 */

export type FileSourceKind = 'browser' | 'wopi' | 'personal';

export type RecentId = string;

export type RecentEntry = {
  id: RecentId;
  name: string;
  sourceFormat: WorkbookFormat | null;
  size: number;
  modifiedAt: number;
};

export type OpenedWorkbook = {
  data: IWorkbookData;
  sourceFormat: WorkbookFormat | null;
  /** Server-backed sources surface the id + etag so the caller can
   *  track them for in-place saves. Browser source leaves these null
   *  — every Save in Mode 1 is a fresh artefact (download or FSA
   *  file write), never an in-place server PUT. */
  serverFileId?: string | null;
  serverEtag?: string | null;
};

export type SaveResult =
  /** Wrote into the user's pinned FSA folder. `folderName` from the
   *  `FileSystemDirectoryHandle`. Toast: "Saved as X in <folder>". */
  | { kind: 'folder'; folderName: string }
  /** Standard download blob — used when no folder is pinned, or the
   *  source is browser-without-FSA. Toast: "Saved as X". */
  | { kind: 'download' }
  /** Wrote to a server host (Mode 3 / Mode 2). `path` is whatever
   *  identifier the source wants the toast to mention. Toast: "Saved
   *  to <path>". `serverEtag` carries the new version so the caller
   *  can update its tracked etag for the next If-Match. */
  | { kind: 'server'; path: string; serverFileId: string; serverEtag: string }
  /** Server returned a stale-version conflict on PUT (HTTP 409 from
   *  WOPI, 412 from /files). The caller (file-actions) opens the
   *  conflict modal — Discard & reload vs Save as copy — and decides
   *  the next move. */
  | { kind: 'conflict'; expectedEtag: string }
  /** User dismissed a native Save-As dialog (desktop bridge). Nothing
   *  was written; the caller shows no "saved" toast. */
  | { kind: 'cancelled' };

export type SaveOptions = {
  filename: string;
  sourceFormat: WorkbookFormat | null;
  /** Set when the workbook on screen was opened from a server-backed
   *  source AND we know its id+etag. When omitted the source falls
   *  back to "save as new" — PersonalFileSource POSTs to /files,
   *  WopiFileSource refuses (its URL token is bound to a single id). */
  existingId?: string | null;
  /** Last-known server etag for `existingId`. Passed as If-Match (or
   *  the WOPI equivalent header) so a stale browser doesn't
   *  overwrite a teammate's save. */
  existingEtag?: string | null;
};

export interface FileSource {
  readonly kind: FileSourceKind;
  /** Human label for UI surfaces ("This browser", "My files", etc.). */
  readonly label: string;

  listRecent(): Promise<RecentEntry[]>;
  /** Subscribe to recent-list changes. Returns the unsubscribe fn. */
  subscribeRecent(cb: () => void): () => void;

  /** Reopen a workbook by id. Throws if `id` is unknown — callers
   *  should treat that as "the user clicked a stale recent". */
  openRecent(id: RecentId): Promise<OpenedWorkbook>;

  /** Drop an entry from the recent list. For the browser source this
   *  removes the IDB record entirely (the only copy lives there); for
   *  server-backed sources this only forgets the recency, not the
   *  file. (The "delete file" affordance there is a separate action.) */
  forgetRecent(id: RecentId): Promise<void>;

  /** Deliver the workbook bytes back to storage. The source decides
   *  the channel (download / FSA / server PUT). */
  save(bytes: Blob, opts: SaveOptions): Promise<SaveResult>;
}
