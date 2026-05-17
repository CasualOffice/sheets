import type { IWorkbookData } from '@univerjs/core';
import { timeItAsync } from '../perf';
import { parseXlsxInWorker } from './parse-in-worker';

/**
 * Public entry point for xlsx import. The actual ExcelJS work lives in a
 * Web Worker (`parser.worker.ts` → `parse-impl.ts`) so the main thread
 * stays responsive while a multi-MB workbook is being parsed. This file
 * stays type-only on the main bundle — ExcelJS doesn't get pulled in
 * here.
 *
 * Fidelity scope (MVP):
 *   - Values + formulas (cell.value / cell.formula)
 *   - Font (family, size, bold, italic, underline, color)
 *   - Fill (solid background)
 *   - Alignment (horizontal, vertical, wrap)
 *   - Number format
 *   - Borders (thin, per side, color preserved)
 *   - Merges
 *   - Sheet order + names
 *
 * Accepts loss: charts, drawings, pivots, validation, conditional formatting,
 * data tables, comments, hyperlinks, advanced borders (dashed/double), themes.
 */

/** A hyperlink read off an xlsx cell that must be replayed into Univer's
 * hyperlink plugin AFTER the snapshot becomes the active workbook. The
 * hyperlink plugin keeps its state in HyperLinkModel (and the rich-text cell
 * body it mutates as a side-effect of AddHyperLinkCommand), neither of which
 * we can construct in the pure-data import path — so we capture the URL
 * here and re-issue the command after the unit mounts.
 *
 * `id` is a fresh client id; xlsx doesn't persist Univer's link ids, so a new
 * one per import is correct and matches what an in-app insert would assign. */
export type PendingHyperlink = {
  subUnitId: string;
  id: string;
  row: number;
  column: number;
  payload: string;
  display?: string;
};

/** Workbook data plus side-channel info that has to be replayed into
 * plugin services after the snapshot is mounted as the active unit. */
export type ImportedWorkbook = IWorkbookData & {
  __pendingHyperlinks?: PendingHyperlink[];
};

export async function xlsxToWorkbookData(buffer: ArrayBuffer): Promise<ImportedWorkbook> {
  return timeItAsync('parse-xlsx', () => parseXlsxInWorker(buffer));
}
