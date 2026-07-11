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

/**
 * Pixel ↔ cell conversion for chart drag / resize. The Univer facade
 * doesn't ship a `pixelToCell` for our coordinate space, so we walk
 * row heights / column widths via `getCellRect` until the cumulative
 * extent contains the target coordinate.
 *
 * Inputs are in canvas-local PRE-scroll coordinates — the same frame
 * `getCellRect` itself returns. Callers must add the viewport scroll
 * before invoking these.
 */

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Sheet = any;

const MAX_SCAN = 2048;

/** Find the row whose [top, bottom) range contains `y`. Clamped to MAX_SCAN. */
export function rowAtY(sheet: Sheet, y: number, hintRow = 0): number {
  if (y < 0) return 0;
  // Walk forward from the hint. Charts only move a few cells per drag
  // frame, so the hint puts us within a handful of iterations.
  let r = Math.max(0, hintRow);
  for (let i = 0; i < MAX_SCAN; i++) {
    let rect: { top: number; bottom: number } | null = null;
    try {
      rect = sheet.getRange(r, 0).getCellRect();
    } catch {
      return Math.max(0, r - 1);
    }
    if (!rect) return Math.max(0, r - 1);
    if (rect.top > y) {
      // Overshot — walk back row by row.
      while (r > 0) {
        r--;
        try {
          const back = sheet.getRange(r, 0).getCellRect();
          if (back && back.top <= y && back.bottom > y) return r;
          if (back && back.top <= y) return r;
        } catch {
          return r;
        }
      }
      return 0;
    }
    if (rect.bottom > y) return r;
    r++;
  }
  return r;
}

/** Find the column whose [left, right) range contains `x`. Clamped to MAX_SCAN. */
export function colAtX(sheet: Sheet, x: number, hintCol = 0): number {
  if (x < 0) return 0;
  let c = Math.max(0, hintCol);
  for (let i = 0; i < MAX_SCAN; i++) {
    let rect: { left: number; right: number } | null = null;
    try {
      rect = sheet.getRange(0, c).getCellRect();
    } catch {
      return Math.max(0, c - 1);
    }
    if (!rect) return Math.max(0, c - 1);
    if (rect.left > x) {
      while (c > 0) {
        c--;
        try {
          const back = sheet.getRange(0, c).getCellRect();
          if (back && back.left <= x && back.right > x) return c;
          if (back && back.left <= x) return c;
        } catch {
          return c;
        }
      }
      return 0;
    }
    if (rect.right > x) return c;
    c++;
  }
  return c;
}

/**
 * Convert a chart's screen-local pixel rect to cell-coordinate position.
 * `screenRect` is the chart's host-local box (left/top/width/height).
 * `host`-local pixels are translated to canvas-local pre-scroll coords
 * using `canvasOffset` (canvas vs. host offset) and `scroll` (the
 * worksheet's current scroll offset). We then snap the resulting
 * top-left and bottom-right cells.
 */
export function rectToCellPos(
  sheet: Sheet,
  screenRect: { left: number; top: number; width: number; height: number },
  canvasOffset: { x: number; y: number },
  scroll: { x: number; y: number },
  hint?: { startRow: number; startColumn: number; endRow: number; endColumn: number },
): { startRow: number; endRow: number; startColumn: number; endColumn: number } | null {
  const preLeft = screenRect.left - canvasOffset.x + scroll.x;
  const preTop = screenRect.top - canvasOffset.y + scroll.y;
  const preRight = preLeft + screenRect.width;
  const preBottom = preTop + screenRect.height;

  const startRow = rowAtY(sheet, preTop, hint?.startRow);
  const startColumn = colAtX(sheet, preLeft, hint?.startColumn);
  // For the bottom-right, snap to the cell BEFORE the bottom-right
  // pixel (so the chart includes that cell, matching Excel's "anchor
  // to bottom-right cell" semantics). Subtract 1 from preRight/Bottom
  // to land in the inclusive last cell.
  const endRow = Math.max(startRow, rowAtY(sheet, preBottom - 1, hint?.endRow ?? startRow));
  const endColumn = Math.max(
    startColumn,
    colAtX(sheet, preRight - 1, hint?.endColumn ?? startColumn),
  );
  return { startRow, endRow, startColumn, endColumn };
}
