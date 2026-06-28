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
 * Sort a range by cell colour — Excel's Data → Sort → Sort On: Cell Color,
 * one level ("<colour> On Top / On Bottom"). Rows whose key-column cell
 * carries the chosen background colour are moved to the top (or bottom) of
 * the selection; the rest keep their relative order. Stable, so re-sorting
 * by a second colour layers predictably.
 *
 * Pure + Univer/React-free so it's unit-testable; the dialog reads the
 * colours via `getBackgrounds()`, the cells via `getCellDatas()`, runs this,
 * and writes back with one atomic `setValues` (preserving value/formula/style).
 */

/** Minimal cell shape we read/write — a subset of Univer's ICellData. */
export interface SortCell {
  v?: unknown;
  f?: string;
  p?: unknown;
  [k: string]: unknown;
}
export type SortRow = Array<SortCell | null>;

export interface ColorSortInput {
  /** Full cell data of the selection, row-major. */
  rows: SortRow[];
  /** Background colour of the key-column cell for each row (same length/order
   *  as `rows`, header included). */
  colors: string[];
  /** Treat the first row as a header (kept in place, never moved). */
  hasHeaders: boolean;
  /** The colour to bring to the top/bottom (compared exactly as read). */
  targetColor: string;
  /** True = matching rows rise to the top; false = sink to the bottom. */
  onTop: boolean;
}

/**
 * Stable-partition the data rows by whether their key colour matches the
 * target. Returns a new matrix (cells cloned) of the same dimensions; the
 * header (if any) stays first.
 */
export function colorSort(input: ColorSortInput): { rows: SortRow[] } {
  const { rows, colors, hasHeaders, targetColor, onTop } = input;
  const headerCount = hasHeaders ? 1 : 0;
  const header = rows.slice(0, headerCount);
  const data = rows.slice(headerCount);
  const dataColors = colors.slice(headerCount);

  const matching: SortRow[] = [];
  const rest: SortRow[] = [];
  data.forEach((row, i) => {
    (dataColors[i] === targetColor ? matching : rest).push(row);
  });

  const ordered = onTop ? [...matching, ...rest] : [...rest, ...matching];

  // Clone every cell: `getCellDatas()` returns live model refs that
  // `setValues` mutates mid-write, so a reordered array of the originals
  // corrupts. Fresh cell objects keep the reorder safe.
  const cloneRow = (row: SortRow): SortRow => row.map((c) => (c == null ? null : { ...c }));
  return { rows: [...header, ...ordered].map(cloneRow) };
}

/** Distinct colours present in a key column, in first-appearance order — the
 *  swatches the dialog offers (header optionally skipped). */
export function distinctColors(colors: string[], hasHeaders: boolean): string[] {
  const seen = new Set<string>();
  const out: string[] = [];
  for (const c of colors.slice(hasHeaders ? 1 : 0)) {
    if (!seen.has(c)) {
      seen.add(c);
      out.push(c);
    }
  }
  return out;
}
