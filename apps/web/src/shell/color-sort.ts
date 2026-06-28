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

export interface ColorMultiSortInput {
  /** Full cell data of the selection, row-major. */
  rows: SortRow[];
  /** Key-column colour per row (same length/order as `rows`, header included). */
  colors: string[];
  /** Treat the first row as a header (kept in place). */
  hasHeaders: boolean;
  /** Colour priority, top-to-bottom. Rows whose colour isn't listed sort
   *  after the listed ones, keeping their original order. */
  order: string[];
}

/**
 * Multi-level colour sort — Excel's Sort dialog with several colour levels
 * ("red on top, then yellow, then green"). Ranks each data row by where its
 * key colour appears in `order`; unlisted colours fall to the end. Stable;
 * returns a new matrix with cells cloned (the `getCellDatas`/`setValues`
 * aliasing guard). The single-colour case is just `order = [colour]`.
 */
export function colorMultiSort(input: ColorMultiSortInput): { rows: SortRow[] } {
  const { rows, colors, hasHeaders, order } = input;
  const headerCount = hasHeaders ? 1 : 0;
  const header = rows.slice(0, headerCount);
  const data = rows.slice(headerCount);
  const dataColors = colors.slice(headerCount);

  const rank = new Map<string, number>();
  order.forEach((c, i) => {
    if (!rank.has(c)) rank.set(c, i);
  });

  const decorated = data.map((row, i) => ({
    row,
    i,
    r: rank.has(dataColors[i]) ? (rank.get(dataColors[i]) as number) : Number.POSITIVE_INFINITY,
  }));
  decorated.sort((a, b) => {
    if (a.r === b.r) return a.i - b.i; // stable within a colour (incl. both unlisted)
    if (a.r === Number.POSITIVE_INFINITY) return 1;
    if (b.r === Number.POSITIVE_INFINITY) return -1;
    return a.r - b.r;
  });

  const cloneRow = (row: SortRow): SortRow => row.map((c) => (c == null ? null : { ...c }));
  return { rows: [...header, ...decorated.map((d) => d.row)].map(cloneRow) };
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
