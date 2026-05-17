import type { IWorkbookData } from '@univerjs/core';
import type { OutlineState } from '../outline/types';
import { timeItAsync } from '../perf';
import { serializeXlsxInWorker } from './serialize-in-worker';

/**
 * Public entry point for xlsx export. The actual ExcelJS work lives in
 * `export-impl.ts` invoked via the `exporter.worker.ts` Web Worker so
 * the main thread stays responsive while a multi-MB save serializes.
 *
 * This file is type-only on the main bundle — ExcelJS doesn't get
 * pulled in here. See `import.ts` for the matching parser-side split
 * and `docs/LARGE_FILE_PIPELINE.md` for the rationale.
 */

/**
 * Cell-level extras the caller has read out of plugin services and wants
 * us to fold into the xlsx output. None of these live on `IWorkbookData`
 * itself, so the export function can't recover them on its own.
 */
export type ExportExtras = {
  /** subUnitId -> rows of { row, column, payload, display } */
  hyperlinks?: Record<string, Array<{ row: number; column: number; payload: string; display?: string }>>;
  /** Per-sheet row/column outline groups — survives the round-trip via two
   *  parallel channels: our `__casual_sheets_outline__` resource (exact
   *  group boundaries) AND ExcelJS row/col `outlineLevel`+`collapsed` (so
   *  Excel renders the native +/- gutter when the file is opened there). */
  outline?: OutlineState;
};

/**
 * Convert a Univer `IWorkbookData` snapshot to an .xlsx Blob.
 * See `import.ts` for the fidelity scope (same coverage in both directions).
 */
export async function workbookDataToXlsx(
  data: IWorkbookData,
  extras: ExportExtras = {},
): Promise<Blob> {
  return timeItAsync('export-xlsx', () => serializeXlsxInWorker(data, extras));
}

// Re-export the constant from its dedicated module so existing imports
// from `./export` keep working.
export { RESOURCES_SHEET } from './constants';
