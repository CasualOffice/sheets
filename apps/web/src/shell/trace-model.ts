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
 * Pure helpers for formula auditing — trace precedents (Excel's Formulas →
 * Trace Precedents). Parses a cell formula's references into cell rectangles
 * so the overlay can draw tracer arrows from each precedent to the cell.
 * Univer/React-free → unit-testable; the layer resolves pixel positions.
 */

import { extractReferences, stripEquals } from './formula-evaluate';

export interface CellRect {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}

export interface RefTarget {
  /** Sheet name if the reference was qualified (`Sheet2!A1`), else null. */
  sheetName: string | null;
  rect: CellRect;
}

/** Column letters → 0-based index (A→0, Z→25, AA→26). */
export function colToIndex(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

function parseCell(s: string): { row: number; col: number } | null {
  const m = /^\$?([A-Za-z]{1,3})\$?([0-9]+)$/.exec(s.trim());
  if (!m) return null;
  const row = Number(m[2]) - 1;
  if (row < 0) return null;
  return { row, col: colToIndex(m[1]) };
}

/**
 * Parse an A1 reference (single cell or range, optionally `$`-absolute or
 * sheet-qualified) into a normalised cell rectangle. Returns null for shapes
 * we don't trace (whole-column `A:A`, names, malformed).
 */
export function parseA1(ref: string): RefTarget | null {
  let sheetName: string | null = null;
  let body = ref.trim();
  const bang = body.lastIndexOf('!');
  if (bang >= 0) {
    sheetName = body.slice(0, bang).replace(/^'(.*)'$/, '$1');
    body = body.slice(bang + 1);
  }
  const [aStr, bStr] = body.split(':');
  const a = parseCell(aStr);
  if (!a) return null;
  const b = bStr ? parseCell(bStr) : a;
  if (!b) return null;
  return {
    sheetName,
    rect: {
      startRow: Math.min(a.row, b.row),
      startCol: Math.min(a.col, b.col),
      endRow: Math.max(a.row, b.row),
      endCol: Math.max(a.col, b.col),
    },
  };
}

/** All references a formula depends on, as cell rectangles (deduped). */
export function precedentsOf(formula: string): RefTarget[] {
  const seen = new Set<string>();
  const out: RefTarget[] = [];
  for (const { sub } of extractReferences(stripEquals(formula))) {
    const t = parseA1(sub);
    if (!t) continue;
    const key = `${t.sheetName ?? ''}!${t.rect.startRow}:${t.rect.startCol}:${t.rect.endRow}:${t.rect.endCol}`;
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(t);
  }
  return out;
}

/** True when a formula references the given target cell — used by trace
 *  dependents to scan candidate cells. */
export function formulaReferences(
  formula: string,
  sheetName: string,
  row: number,
  col: number,
): boolean {
  for (const t of precedentsOf(formula)) {
    // Unqualified refs resolve to the formula's own sheet (the caller passes
    // the candidate cell's sheet as `sheetName`).
    if (t.sheetName != null && t.sheetName !== sheetName) continue;
    const { rect } = t;
    if (row >= rect.startRow && row <= rect.endRow && col >= rect.startCol && col <= rect.endCol) {
      return true;
    }
  }
  return false;
}
