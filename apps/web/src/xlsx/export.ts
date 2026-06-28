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

import type { IWorkbookData } from '@univerjs/core';
import type { OutlineState } from '../outline/types';
import type { ChartModel } from '../charts/types';
import type { PivotModel } from '../pivots/types';
import type { SparklineModel } from '../sparklines/types';
import { timeItAsync } from '../perf';
import { workbookDataToXlsx as sdkWorkbookDataToXlsx } from '@casualoffice/sheets/xlsx';
import { writeOutlineIntoSnapshot } from '../outline/resources';
import { writeChartsIntoSnapshot } from '../charts/resources';
import { writePivotsIntoSnapshot } from '../pivots/resources';
import { writeSparklinesIntoSnapshot } from '../sparklines/resources';
import { writeWatchesIntoSnapshot } from '../shell/watch-resources';
import type { Watch } from '../shell/watch-model';
import { injectNativePivots } from './pivot-native-export';

/**
 * Public entry point for xlsx export. The core converter now lives in the SDK
 * (`@casualoffice/sheets/xlsx`, runs in its own Web Worker so a multi-MB save
 * stays off the main thread + keeps ExcelJS out of this bundle). This file is
 * the app's power-host layer: it bakes app-only feature models (charts, pivots,
 * sparklines, outline) into the snapshot's resources, then delegates the
 * cell/style/merge/resource serialization to the shared SDK exporter — so the
 * app and third-party SDK hosts share one exporter. See `import.ts` for the
 * matching parser side and `docs/LARGE_FILE_PIPELINE.md` for the rationale.
 */

/**
 * Cell-level extras the caller has read out of plugin services and wants
 * us to fold into the xlsx output. None of these live on `IWorkbookData`
 * itself, so the export function can't recover them on its own.
 */
export type ExportExtras = {
  /** subUnitId -> rows of { row, column, payload, display } */
  hyperlinks?: Record<
    string,
    Array<{ row: number; column: number; payload: string; display?: string }>
  >;
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
  /** Watch Window cells stashed under `__casual_sheets_watches__`. Same
   *  pattern as pivots/charts — survives our round-trip; foreign readers
   *  ignore the resource. */
  watches?: Watch[];
};

/**
 * Convert a Univer `IWorkbookData` snapshot to an .xlsx Blob.
 * See `import.ts` for the fidelity scope (same coverage in both directions).
 */
export async function workbookDataToXlsx(
  data: IWorkbookData,
  extras: ExportExtras = {},
): Promise<Blob> {
  // Bake app-only feature models into the snapshot resources before handing it
  // to the shared SDK exporter. Done on a shallow clone (custom + resources are
  // the only things the writers touch — both reassigned, not deep-mutated) so a
  // caller that reuses its snapshot doesn't see the export-side additions. The
  // big `sheets`/cellData refs are shared, not copied.
  const baked: IWorkbookData = {
    ...data,
    custom: { ...data.custom },
    resources: [...(data.resources ?? [])],
  };
  if (extras.outline && Object.keys(extras.outline).length > 0) {
    writeOutlineIntoSnapshot(baked, extras.outline);
  }
  if (extras.charts && extras.charts.length > 0) {
    writeChartsIntoSnapshot(baked, extras.charts);
  }
  if (extras.pivots && extras.pivots.length > 0) {
    writePivotsIntoSnapshot(baked, extras.pivots);
  }
  if (extras.sparklines && extras.sparklines.length > 0) {
    writeSparklinesIntoSnapshot(baked, extras.sparklines);
  }
  if (extras.watches && extras.watches.length > 0) {
    writeWatchesIntoSnapshot(baked, extras.watches);
  }
  // The SDK exporter handles the generic xlsx-native bits it can derive on its
  // own (hyperlink cells, outline gutter, floating chart images).
  const blob = await timeItAsync('export-xlsx', () =>
    sdkWorkbookDataToXlsx(baked, {
      hyperlinks: extras.hyperlinks,
      outline: extras.outline,
      chartImages: extras.chartImages,
    }),
  );
  // Opt-in (off by default): also emit native xl/pivotTables for in-app pivots
  // so Excel re-opens them as real PivotTables. No-op + safe fallback when off.
  return injectNativePivots(blob, baked, extras.pivots);
}

// Re-export the constant from its dedicated module so existing imports
// from `./export` keep working.
export { RESOURCES_SHEET } from '@casualoffice/sheets/xlsx';
