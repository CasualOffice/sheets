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
import { timeItAsync } from './_perf';
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

/**
 * Workbook data ready to mount. Stage 5 of the pipeline folded
 * hyperlinks into `cell.p.body.customRanges` inline, so no more
 * `__pendingHyperlinks` side-channel — the snapshot is self-contained.
 */
export type ImportedWorkbook = IWorkbookData;

export async function xlsxToWorkbookData(buffer: ArrayBuffer): Promise<ImportedWorkbook> {
  return timeItAsync('parse-xlsx', () => parseXlsxInWorker(buffer));
}
