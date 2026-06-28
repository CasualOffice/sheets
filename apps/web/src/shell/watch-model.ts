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
 * Pure model for the Watch Window (Excel's Formulas → Watch Window). A watch
 * pins a single cell so its value + formula stay visible in the panel even
 * when you scroll away or switch sheets. Kept Univer/React-free so the list
 * operations are unit-testable; the panel resolves live value/formula.
 */

export interface Watch {
  id: string;
  sheetId: string;
  sheetName: string;
  row: number;
  col: number;
}

/** Stable identity of a watched cell (a cell is watched at most once). */
export function watchKey(sheetId: string, row: number, col: number): string {
  return `${sheetId}!${row}:${col}`;
}

/** A1-style cell reference (no sheet qualifier). */
export function cellA1(row: number, col: number): string {
  let n = col;
  let letters = '';
  do {
    letters = String.fromCharCode(65 + (n % 26)) + letters;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return `${letters}${row + 1}`;
}

/**
 * Append new watches, skipping cells already watched (dedup by sheet+cell).
 * `id` is supplied by the caller so the pure function stays deterministic.
 */
export function addWatches(
  list: Watch[],
  additions: Array<Omit<Watch, 'id'>>,
  makeId: (index: number) => string,
): Watch[] {
  const seen = new Set(list.map((w) => watchKey(w.sheetId, w.row, w.col)));
  const next = [...list];
  let added = 0;
  for (const a of additions) {
    const key = watchKey(a.sheetId, a.row, a.col);
    if (seen.has(key)) continue;
    seen.add(key);
    next.push({ ...a, id: makeId(added) });
    added += 1;
  }
  return next;
}

/** Remove a watch by id. */
export function removeWatch(list: Watch[], id: string): Watch[] {
  return list.filter((w) => w.id !== id);
}

/** Enumerate the cells of a rectangle, capped so a giant selection can't
 *  flood the window (Excel caps watches too). */
export function cellsInRect(
  rect: { startRow: number; endRow: number; startColumn: number; endColumn: number },
  cap = 100,
): Array<{ row: number; col: number }> {
  const out: Array<{ row: number; col: number }> = [];
  for (let r = rect.startRow; r <= rect.endRow; r++) {
    for (let c = rect.startColumn; c <= rect.endColumn; c++) {
      out.push({ row: r, col: c });
      if (out.length >= cap) return out;
    }
  }
  return out;
}
