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

/**
 * Formula-reference navigation — Excel's Ctrl+[ (precedents) and
 * Ctrl+] (dependents).
 *
 *   - **Ctrl+[**: select every cell the *current* cell's formula
 *     references.
 *   - **Ctrl+]**: select every cell whose formula references the
 *     *current* cell.
 *
 * Parsing is regex-based — handles the common cases (`A1`, `$A$1`,
 * `A1:B5`, `Sheet1!A1`) but skips named ranges, complex expressions,
 * and cross-workbook refs. Those are rare enough that the heuristic
 * version is the right v1; the alternative is wiring through Univer's
 * formula AST which doubles the surface area.
 */

const REF_RE = /(?:([A-Za-z_][\w]*)!)?(\$?[A-Z]+\$?\d+)(?::(\$?[A-Z]+\$?\d+))?/g;

type CellRef = { sheet: string | null; sr: number; sc: number; er: number; ec: number };

function colLettersToIndex(letters: string): number {
  const stripped = letters.replace(/\$/g, '');
  let n = 0;
  for (let i = 0; i < stripped.length; i += 1) {
    n = n * 26 + (stripped.charCodeAt(i) - 64);
  }
  return n - 1;
}

function parseA1(a1: string): { row: number; col: number } | null {
  const m = /^\$?([A-Z]+)\$?(\d+)$/.exec(a1);
  if (!m) return null;
  return { col: colLettersToIndex(m[1]), row: parseInt(m[2], 10) - 1 };
}

export function extractRefs(formula: string): CellRef[] {
  // Strip the leading '=' so the regex can scan content uniformly.
  const body = formula.startsWith('=') ? formula.slice(1) : formula;
  const refs: CellRef[] = [];
  REF_RE.lastIndex = 0;
  let m: RegExpExecArray | null;
  // eslint-disable-next-line no-cond-assign
  while ((m = REF_RE.exec(body)) !== null) {
    const sheet = m[1] ?? null;
    const start = parseA1(m[2]);
    if (!start) continue;
    const end = m[3] ? parseA1(m[3]) : start;
    if (!end) continue;
    refs.push({
      sheet,
      sr: Math.min(start.row, end.row),
      er: Math.max(start.row, end.row),
      sc: Math.min(start.col, end.col),
      ec: Math.max(start.col, end.col),
    });
  }
  return refs;
}

function activeCell(api: FUniver): {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  wb: any;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  ws: any;
  row: number;
  col: number;
} | null {
  const wb = api.getActiveWorkbook();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const ws = wb?.getActiveSheet() as any;
  if (!wb || !ws) return null;
  const range = ws.getActiveRange?.();
  if (!range) return null;
  return { wb, ws, row: range.getRow(), col: range.getColumn() };
}

function setSelectionRanges(
  api: FUniver,
  ranges: Array<{ sr: number; er: number; sc: number; ec: number }>,
): void {
  if (ranges.length === 0) return;
  const wb = api.getActiveWorkbook();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const ws = wb?.getActiveSheet() as any;
  if (!wb || !ws) return;
  // Activate the first range — keeps focus rules sane. v1 doesn't
  // build a Ctrl-click multi-range selection (Univer's facade for
  // that requires a deeper hook into SheetsSelectionsService); a
  // single combined bounding box is the practical compromise.
  let sr = ranges[0].sr;
  let sc = ranges[0].sc;
  let er = ranges[0].er;
  let ec = ranges[0].ec;
  for (const r of ranges) {
    sr = Math.min(sr, r.sr);
    sc = Math.min(sc, r.sc);
    er = Math.max(er, r.er);
    ec = Math.max(ec, r.ec);
  }
  try {
    ws.getRange(sr, sc, er - sr + 1, ec - sc + 1).activate();
  } catch (err) {
    console.warn('[formula-refs] failed to activate range', err);
  }
}

/**
 * Ctrl+[ — select every cell the active cell's formula references.
 * Same-sheet refs only in v1 (cross-sheet would need a sheet-tab swap
 * + range activation; tracked as a follow-up).
 */
export function selectPrecedents(api: FUniver): void {
  const at = activeCell(api);
  if (!at) return;
  const data = at.ws.getRange(at.row, at.col).getCellData();
  if (!data?.f) return;
  const refs = extractRefs(String(data.f)).filter((r) => r.sheet === null);
  setSelectionRanges(api, refs);
}

/**
 * Ctrl+] — walk every cell in the active sheet, parse its formula,
 * and collect those that reference the current cell.
 *
 * O(cells with formulas) per invocation. Caches nothing — Excel's own
 * dependency graph is incremental but ours uses Univer's formula
 * engine results indirectly. Fine for a few-thousand-formula
 * workbook; a future optimisation could subscribe to the engine's
 * adjacency map.
 */
export function selectDependents(api: FUniver): void {
  const at = activeCell(api);
  if (!at) return;
  const targetRow = at.row;
  const targetCol = at.col;
  const dependents: Array<{ sr: number; er: number; sc: number; ec: number }> = [];
  // Walk the used range so we don't iterate 100k empty cells.
  const used = at.ws.getDataRange?.() ?? at.ws.getActiveRange?.();
  if (!used) return;
  const startR = used.getRow();
  const startC = used.getColumn();
  const endR = startR + used.getHeight() - 1;
  const endC = startC + used.getWidth() - 1;
  for (let r = startR; r <= endR; r += 1) {
    for (let c = startC; c <= endC; c += 1) {
      const d = at.ws.getRange(r, c).getCellData();
      const f = d?.f;
      if (!f) continue;
      const refs = extractRefs(String(f));
      const hit = refs.some(
        (ref) =>
          ref.sheet === null &&
          targetRow >= ref.sr &&
          targetRow <= ref.er &&
          targetCol >= ref.sc &&
          targetCol <= ref.ec,
      );
      if (hit) dependents.push({ sr: r, er: r, sc: c, ec: c });
    }
  }
  setSelectionRanges(api, dependents);
}
