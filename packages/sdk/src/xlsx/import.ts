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
 * Public entry point for xlsx import. Browser main threads dispatch to a
 * Web Worker (`parser.worker.ts` → `parse-impl.ts`) so a multi-MB workbook
 * does not block the UI. Node and existing Worker contexts run the same
 * converter directly; that implementation loads ExcelJS only when invoked.
 *
 * Fidelity scope:
 *   - Values + formulas (cell.value / cell.formula)
 *   - Font (family, size, bold, italic, underline, color)
 *   - Fill (solid background)
 *   - Alignment (horizontal, vertical, wrap)
 *   - Number format
 *   - Borders
 *   - Merges
 *   - Sheet order + names
 *   - Tables, comments, hyperlinks, data validation, conditional formatting,
 *     page setup, named ranges, and opaque Univer resources
 *   - Raw OOXML passthrough for drawings, pivots, macros, external links, and
 *     threaded comments that Univer or ExcelJS cannot model losslessly
 *
 * Passthrough features survive save/download but are not necessarily editable
 * or rendered inside Univer (for example VBA and pivot machinery).
 */

/**
 * Workbook data ready to mount. Stage 5 of the pipeline folded
 * hyperlinks into `cell.p.body.customRanges` inline, so no more
 * `__pendingHyperlinks` side-channel — the snapshot is self-contained.
 */
export type ImportedWorkbook = IWorkbookData;

export async function xlsxToWorkbookData(buffer: ArrayBuffer): Promise<ImportedWorkbook> {
  return timeItAsync('parse-xlsx', async () => {
    // Browser main threads keep the responsive Worker path. Node and existing
    // Worker contexts have no `window`, so run the same pure converter directly
    // instead of requiring browser Worker globals.
    if (typeof window === 'undefined' || typeof Worker === 'undefined') {
      const { workbookFromExcelJs } = await import('./parse-impl');
      return workbookFromExcelJs(buffer);
    }
    return parseXlsxInWorker(buffer);
  });
}
