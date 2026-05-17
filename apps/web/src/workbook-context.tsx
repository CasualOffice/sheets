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
   */
  replaceWorkbook: (next: IWorkbookData, sourceFormat?: WorkbookFormat | null) => void;
  /**
   * Rename the active workbook in place — updates `meta.name` and calls
   * Univer's `setName` on the active workbook. No re-mount.
   */
  renameWorkbook: (name: string) => void;
};

export const WorkbookContext = createContext<WorkbookCtxValue | null>(null);
