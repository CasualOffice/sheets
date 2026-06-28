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
 * Subtotals — pure layout logic (Excel's Data → Subtotal).
 *
 * Given the selected list (header + data rows), insert a subtotal row at each
 * change in the grouping column and a grand-total row at the end. Every inserted
 * row uses `SUBTOTAL(code, range)` — which ignores nested SUBTOTALs — so the
 * grand total can span the whole list without double-counting and each formula
 * is just an A1 range over that group's final row positions.
 *
 * Kept free of Univer/React so it's unit-testable; the dialog feeds it
 * `getCellDatas()` + the absolute origin of the selection, inserts the right
 * number of rows, and writes the result back with one `setValues`.
 */

export interface SubtotalCell {
  v?: unknown;
  f?: string;
  p?: unknown;
  [k: string]: unknown;
}

export type SubtotalRow = Array<SubtotalCell | null>;

export interface SubtotalInput {
  /** Full selection cell data incl. the header row. */
  rows: SubtotalRow[];
  /** Column index (within the selection) whose changes start a new group. */
  groupCol: number;
  /** SUBTOTAL function code (9=SUM, 1=AVERAGE, 2=COUNT, 3=COUNTA, 4=MAX, 5=MIN, 6=PRODUCT). */
  functionCode: number;
  /** Column indices (within the selection) to subtotal. */
  subtotalCols: number[];
  /** Absolute sheet row of the selection's top-left (0-based). */
  startRow: number;
  /** Absolute sheet column of the selection's top-left (0-based). */
  startCol: number;
}

export interface SubtotalResult {
  /** New block: header + data interleaved with subtotal rows + grand total. */
  rows: SubtotalRow[];
  /** Rows added (rows.length − input rows.length). */
  insertedRows: number;
  /** Number of group subtotal rows added (excludes the grand total). */
  groups: number;
}

function colToLetters(col: number): string {
  let s = '';
  let n = col;
  do {
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return s;
}

function valueOf(cell: SubtotalCell | null): string {
  if (cell == null) return '';
  if (cell.v != null) return String(cell.v);
  const ds = (cell.p as { body?: { dataStream?: string } } | undefined)?.body?.dataStream;
  if (typeof ds === 'string') return ds.replace(/[\r\n]+$/, '');
  return '';
}

export function computeSubtotals(input: SubtotalInput): SubtotalResult {
  const { rows, groupCol, functionCode, subtotalCols, startRow, startCol } = input;
  const width = rows[0]?.length ?? 0;
  const header = rows.slice(0, 1);
  const data = rows.slice(1);
  if (data.length === 0 || width === 0) {
    return { rows, insertedRows: 0, groups: 0 };
  }

  // Group consecutive data rows by the grouping column's value.
  const groups: { key: string; rows: SubtotalRow[] }[] = [];
  for (const row of data) {
    const key = valueOf(row[groupCol] ?? null);
    const last = groups[groups.length - 1];
    if (last && last.key === key) last.rows.push(row);
    else groups.push({ key, rows: [row] });
  }

  const emptyRow = (): SubtotalRow => new Array(width).fill(null);
  const subtotalCell = (col: number, firstAbsRow: number, lastAbsRow: number): SubtotalCell => {
    const letter = colToLetters(startCol + col);
    return {
      f: `=SUBTOTAL(${functionCode},${letter}${firstAbsRow + 1}:${letter}${lastAbsRow + 1})`,
    };
  };

  // Deep-ish copy so the result never aliases the live cell matrix the caller
  // read with getCellDatas() (writing it back could otherwise self-overwrite).
  const cloneRow = (r: SubtotalRow): SubtotalRow => r.map((c) => (c == null ? null : { ...c }));

  // Build the output sequence, tracking absolute row positions as we go.
  const out: SubtotalRow[] = header.map(cloneRow);
  let abs = startRow + out.length; // absolute row of the next output row

  for (const group of groups) {
    const firstDataAbs = abs;
    for (const r of group.rows) {
      out.push(cloneRow(r));
      abs += 1;
    }
    const lastDataAbs = abs - 1;
    // Subtotal row for this group.
    const row = emptyRow();
    row[groupCol] = { v: `${group.key} Total` };
    for (const col of subtotalCols) {
      if (col === groupCol) continue; // label wins on the group column
      row[col] = subtotalCell(col, firstDataAbs, lastDataAbs);
    }
    out.push(row);
    abs += 1;
  }

  // Grand total over the whole list (SUBTOTAL ignores the nested subtotals).
  const firstAllAbs = startRow + header.length;
  const lastAllAbs = abs - 1;
  const grand = emptyRow();
  grand[groupCol] = { v: 'Grand Total' };
  for (const col of subtotalCols) {
    if (col === groupCol) continue;
    grand[col] = subtotalCell(col, firstAllAbs, lastAllAbs);
  }
  out.push(grand);

  return { rows: out, insertedRows: out.length - rows.length, groups: groups.length };
}
