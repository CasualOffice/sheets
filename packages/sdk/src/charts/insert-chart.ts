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
import type { ChartType, ChartModel } from './types';
import { newChartId } from './types';

type Range = { startRow: number; endRow: number; startColumn: number; endColumn: number };

/**
 * Build a ChartModel for `source` on the active sheet. The chart anchors
 * to a 10-row × 8-col block placed two rows below the source so it doesn't
 * cover its own data. Mirrors Excel's Insert &gt; Chart defaults — source =
 * the dialog's selection, output = roughly the right size, dropped below.
 *
 * Returns null when there's no active workbook/sheet, or the range has
 * fewer than 2 rows × 2 cols (no header + data row, or no label + value
 * column).
 */
export function buildChartModelForRange(
  api: FUniver,
  source: Range,
  type: ChartType,
): ChartModel | null {
  const wb = api.getActiveWorkbook();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const ws = wb?.getActiveSheet() as any;
  if (!wb || !ws) return null;

  const rows = source.endRow - source.startRow + 1;
  const cols = source.endColumn - source.startColumn + 1;
  if (rows < 2 || cols < 2) return null;

  const chartTop = source.endRow + 2;
  const chartLeft = source.startColumn;
  return {
    id: newChartId(),
    sheetId: ws.getSheetId(),
    source,
    pos: {
      startRow: chartTop,
      endRow: chartTop + 9,
      startColumn: chartLeft,
      endColumn: chartLeft + 7,
    },
    type,
  };
}

/**
 * Read the active selection on the active sheet. Returns null if no
 * workbook / sheet / selection is available — the menu item is enabled
 * unconditionally, so the caller still has to handle the empty case.
 */
export function getActiveSelectionRange(api: FUniver): Range | null {
  const wb = api.getActiveWorkbook();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const ws = wb?.getActiveSheet() as any;
  const range = ws?.getActiveRange();
  if (!wb || !ws || !range) return null;
  const r = range.getRange();
  return {
    startRow: r.startRow,
    endRow: r.endRow,
    startColumn: r.startColumn,
    endColumn: r.endColumn,
  };
}

/**
 * Format a cell range as an A1-style reference (e.g. `A1:C4`). Used by
 * the insert dialog to pre-fill its source-range input from the current
 * selection so the common case is two clicks.
 */
export function rangeToA1(range: Range): string {
  const tl = `${colIndexToA1(range.startColumn)}${range.startRow + 1}`;
  if (range.startRow === range.endRow && range.startColumn === range.endColumn) {
    return tl;
  }
  const br = `${colIndexToA1(range.endColumn)}${range.endRow + 1}`;
  return `${tl}:${br}`;
}

function colIndexToA1(c: number): string {
  let n = c + 1;
  let s = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
