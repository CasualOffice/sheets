import type { IWorkbookData } from '@univerjs/core';
import type { OutlineState } from '../outline/types';
import type { ChartModel } from '../charts/types';
import type { PivotModel } from '../pivots/types';
import type { SparklineModel } from '../sparklines/types';
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
  /** Chart models stashed under `__casual_sheets_charts__`. Survives our
   *  round-trip with full editability. To make charts *visible* in Excel
   *  itself, Charts P5b also passes pre-rendered PNG snapshots via
   *  `chartImages` below — Excel renders them as static images while our
   *  app re-attaches the live chart from the JSON sidecar on re-open. */
  charts?: ChartModel[];
  /** Pre-rendered chart bitmaps to embed in the xlsx as floating images,
   *  one per chart. Anchored to the same cell rectangle as the source
   *  ChartModel.pos. The main thread renders these via ECharts before
   *  posting to the exporter worker (ECharts needs a DOM). */
  chartImages?: Array<{
    chartId: string;
    sheetId: string;
    png: ArrayBuffer;
    anchor: { startRow: number; endRow: number; startColumn: number; endColumn: number };
  }>;
  /** Pivot models stashed under `__casual_sheets_pivots__`. The computed
   *  pivot output lives in regular cell values (so it round-trips through
   *  any xlsx reader); this carries the *definition* so a future refresh
   *  / change-source action can re-run apply without the user
   *  reconfiguring. */
  pivots?: PivotModel[];
  /** Sparkline models stashed under `__casual_sheets_sparklines__`.
   *  Same pattern as charts/pivots — survives our round-trip, lost in
   *  foreign xlsx readers (Excel renders the underlying source cells
   *  but not the mini-chart overlay). */
  sparklines?: SparklineModel[];
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
export { RESOURCES_SHEET } from '@schnsrw/casual-sheets/xlsx';
