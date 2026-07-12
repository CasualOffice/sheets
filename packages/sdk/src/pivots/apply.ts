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

import type { FUniver } from '@univerjs/core/facade';
import { computePivot, type PivotGrid, type SourceMatrix } from './compute';
import type { PivotModel } from './types';

/**
 * Read the source range, compute the pivot, then write the resulting
 * cell grid into the target sheet starting at `model.target`. The
 * write is one `setRangeValues` call so it lands as a single Univer
 * mutation — collab and undo both treat it atomically.
 *
 * If `prevExtent` is supplied (refresh path), the previous output
 * rectangle is cleared first so a shrunk pivot doesn't leave residual
 * rows from the prior write. Insert paths pass null.
 */
export function applyPivot(
  api: FUniver,
  model: PivotModel,
  prevExtent?: { rows: number; cols: number } | null,
): { rows: number; cols: number } | null {
  const wb = api.getActiveWorkbook();
  if (!wb) return null;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const sheets = wb.getSheets() as any[];
  const sourceWs = sheets.find((s) => s.getSheetId?.() === model.sourceSheetId);
  const targetWs = sheets.find((s) => s.getSheetId?.() === model.targetSheetId);
  if (!sourceWs || !targetWs) return null;

  const matrix = readSourceMatrix(sourceWs, model.source);
  if (matrix.records.length === 0) return null;

  const { grid } = computePivot(matrix, model);
  if (grid.length === 0) return null;

  // Clear the previous extent first if we have one — keeps stale rows
  // from leaking through when the new grid is smaller (e.g. a filter
  // narrowed the row keys, or the source range shrank).
  if (prevExtent && (prevExtent.rows > grid.length || prevExtent.cols > (grid[0]?.length ?? 0))) {
    clearGridArea(targetWs, model.target, prevExtent);
  }

  writeGridToSheet(targetWs, model.target, grid);
  return { rows: grid.length, cols: grid[0]?.length ?? 0 };
}

/**
 * P1 — re-read source data and re-apply the pivot. Refreshing is what
 * makes pivots track upstream edits; without it the output frozen at
 * insert time goes stale. Returns the new extent (or null on failure).
 */
export function refreshPivot(
  api: FUniver,
  model: PivotModel,
): { rows: number; cols: number } | null {
  return applyPivot(api, model, model.lastOutputExtent ?? null);
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function readSourceMatrix(ws: any, src: PivotModel['source']): SourceMatrix {
  const headers: string[] = [];
  for (let c = src.startColumn; c <= src.endColumn; c++) {
    const v = ws.getRange(src.startRow, c).getValue();
    headers.push(v == null ? '' : String(v));
  }
  const records: Array<Array<string | number | null>> = [];
  for (let r = src.startRow + 1; r <= src.endRow; r++) {
    const row: Array<string | number | null> = [];
    let anyValue = false;
    for (let c = src.startColumn; c <= src.endColumn; c++) {
      const v = ws.getRange(r, c).getValue();
      if (v == null || v === '') {
        row.push(null);
      } else if (typeof v === 'number' || typeof v === 'string') {
        row.push(v);
        anyValue = true;
      } else if (typeof v === 'boolean') {
        row.push(v ? 1 : 0);
        anyValue = true;
      } else {
        row.push(String(v));
        anyValue = true;
      }
    }
    // Skip blank rows — Excel treats them as terminators, but a
    // misclick that includes an extra blank row in the selection
    // shouldn't introduce a phantom "(blank)" key. Drop them.
    if (anyValue) records.push(row);
  }
  return { headers, records };
}

/** Blank an arbitrary rectangle on the sheet — used by refresh to
 *  reset the previous output before writing the new grid. */
function clearGridArea(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  ws: any,
  target: { row: number; column: number },
  extent: { rows: number; cols: number },
): void {
  const blank: Array<Array<{ v: null }>> = [];
  for (let r = 0; r < extent.rows; r += 1) {
    const row: Array<{ v: null }> = [];
    for (let c = 0; c < extent.cols; c += 1) {
      row.push({ v: null });
    }
    blank.push(row);
  }
  const range = ws.getRange(target.row, target.column, extent.rows, extent.cols);
  range.setValues(blank);
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function writeGridToSheet(ws: any, target: { row: number; column: number }, grid: PivotGrid): void {
  const rows = grid.length;
  const cols = grid[0]?.length ?? 0;
  if (rows === 0 || cols === 0) return;
  // Build the IRange-shaped object setRangeValues expects: a 2D
  // array of `{ v }` cell objects keyed by `[r-offset][c-offset]`.
  const cellMatrix: Array<Array<{ v: string | number | null }>> = grid.map((rowVals) =>
    rowVals.map((v) => ({ v })),
  );
  const range = ws.getRange(target.row, target.column, rows, cols);
  range.setValues(cellMatrix);
}
