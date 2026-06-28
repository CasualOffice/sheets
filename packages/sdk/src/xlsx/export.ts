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
import { serializeXlsxInWorker } from './serialize-in-worker';

/**
 * xlsx EXPORT — Univer `IWorkbookData` snapshot → `.xlsx` Blob.
 *
 * The core converter (cells, formulas, styles, merges, number formats,
 * borders, hyperlinks, comments, data validation, tables, page setup,
 * named ranges, VBA passthrough) lives here in the SDK and runs in a Web
 * Worker so a multi-MB save doesn't block the main thread. Fidelity is the
 * mirror of the importer (`./import`).
 *
 * Anything NOT recoverable from the snapshot alone is passed via `ExportExtras`:
 * live hyperlink cells read from the plugin, native outline gutter levels, and
 * pre-rendered floating images (e.g. chart bitmaps). These are deliberately
 * GENERIC shapes — the SDK exporter knows nothing about charts/pivots/sparkline
 * *models*; a power host (apps/web) bakes those into the snapshot's resources
 * itself before calling this.
 */
export type ExportExtras = {
  /** subUnitId → cells to render as xlsx-native `{ text, hyperlink }`. */
  hyperlinks?: Record<
    string,
    Array<{ row: number; column: number; payload: string; display?: string }>
  >;
  /** subUnitId → row/column outline groups → xlsx-native `outlineLevel` gutter. */
  outline?: Record<
    string,
    {
      rows?: Array<{ start: number; end: number; collapsed?: boolean }>;
      cols?: Array<{ start: number; end: number; collapsed?: boolean }>;
    }
  >;
  /** Pre-rendered bitmaps embedded as floating images, anchored to a cell
   *  rectangle. The SDK doesn't render these — the host supplies the PNG. */
  chartImages?: Array<{
    chartId: string;
    sheetId: string;
    png: ArrayBuffer;
    anchor: { startRow: number; endRow: number; startColumn: number; endColumn: number };
  }>;
};

/**
 * Convert a Univer `IWorkbookData` snapshot to an `.xlsx` Blob. See `./import`
 * for the matching fidelity scope (same coverage in both directions).
 */
export async function workbookDataToXlsx(
  data: IWorkbookData,
  extras: ExportExtras = {},
): Promise<Blob> {
  return serializeXlsxInWorker(data, extras);
}
