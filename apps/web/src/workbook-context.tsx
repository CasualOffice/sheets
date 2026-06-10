import { createContext, type RefObject } from 'react';
import type { IWorkbookData } from '@univerjs/core';

/**
 * The format the current workbook was loaded from. Drives the Save action so
 * Ctrl+S writes back in the same format the user opened (Excel and LibreOffice
 * both behave this way). `null` means "started blank" — Save defaults to xlsx.
 */
export type WorkbookFormat = 'xlsx' | 'ods' | 'csv' | 'tsv';

/**
 * Small react-state shape carried per active workbook. The full
 * `IWorkbookData` snapshot does NOT live in React state — only this
 * metadata does. That's pipeline Stage 3: the prior design kept two
 * copies of the entire workbook in memory (React state + Univer's
 * internal copy), which doubled the peak heap on a 100 MB file.
 *
 * `revision` is bumped on every `replaceWorkbook` call so effects can
 * subscribe to "workbook was swapped" without depending on snapshot
 * identity.
 */
export type WorkbookMeta = {
  id: string;
  name: string;
  sourceFormat: WorkbookFormat | null;
  revision: number;
  /** Server-backed file id when the workbook was opened via
   *  PersonalFileSource.openRecent / WopiFileSource.openRecent. Null
   *  for browser-source workbooks (download / FSA / templates). The
   *  Save flow passes this back as `existingId` so the server can
   *  PUT in place instead of creating a duplicate entry. */
  serverFileId?: string | null;
  /** Last-known server etag for `serverFileId`. Sent as If-Match
   *  (or `X-WOPI-ItemVersion`) on the next save so a stale browser
   *  doesn't overwrite a teammate's edit. */
  serverEtag?: string | null;
};

/**
 * Active version-history preview, if any. Non-null while the user is
 * examining a past snapshot in the main grid; the PreviewBanner
 * renders from this shape and offers Restore / Cancel.
 */
export type PreviewState = {
  versionId: number;
  versionName: string;
  versionSavedAt: number;
};

export type WorkbookCtxValue = {
  meta: WorkbookMeta;
  /**
   * Mounting/swap callers (UniverSheet, OutlineProvider) read the full
   * snapshot from here during the revision-change effect. The ref's
   * `.current` is set right before `replaceWorkbook` notifies React,
   * and cleared a tick later so the workbook becomes GC-eligible after
   * mount.
   */
  snapshotRef: RefObject<IWorkbookData | null>;
  /**
   * Swap the active workbook. Stores `next` on `snapshotRef.current`,
   * bumps `meta.revision`, and schedules the ref to be cleared so the
   * snapshot frees once React has finished its render pass.
   * `server` carries the file id + etag when the source is server-
   * backed (PersonalFileSource / WopiFileSource); omit for browser-
   * source flows.
   */
  replaceWorkbook: (
    next: IWorkbookData,
    sourceFormat?: WorkbookFormat | null,
    server?: { fileId: string | null; etag: string | null } | null,
  ) => void;
  /** Update the tracked server etag after a successful in-place
   *  save — the file id stays the same; only the version moves. */
  updateServerEtag: (etag: string | null) => void;
  /** Bind a `serverFileId` after the FIRST successful create-save of
   *  a workbook that started life as a draft (route `/sheet/new` or a
   *  blank New). Subsequent saves take the in-place PUT path instead
   *  of creating duplicate rows (UX_AUDIT.md §2.3). Also rewrites
   *  the browser URL from `/sheet/new` → `/sheet/<id>` via
   *  history.replaceState so back / refresh / bookmark all converge. */
  updateServerFileId: (fileId: string | null) => void;
  /**
   * Rename the active workbook in place — updates `meta.name` and calls
   * Univer's `setName` on the active workbook. No re-mount.
   */
  renameWorkbook: (name: string) => void;
  /** Version-history preview state. Null when not previewing. */
  preview: PreviewState | null;
  /**
   * Enter preview: caller supplies the live workbook state (read from
   * `wb.save()` before calling) so App can stash it for the cancel
   * path. App then swaps to the snapshot via replaceWorkbook.
   */
  enterPreview: (
    versionId: number,
    versionName: string,
    versionSavedAt: number,
    snapshotData: IWorkbookData,
    snapshotSourceFormat: WorkbookFormat | null,
    currentLiveData: IWorkbookData,
    currentLiveFormat: WorkbookFormat | null,
  ) => void;
  /** Leave preview, restoring the pre-preview workbook. */
  exitPreview: () => void;
  /** Make the previewed snapshot the live workbook. The caller is
   *  responsible for capturing the pre-restore live state as a
   *  manual version FIRST (so undo of restore exists). */
  commitPreview: () => void;
};

export const WorkbookContext = createContext<WorkbookCtxValue | null>(null);
