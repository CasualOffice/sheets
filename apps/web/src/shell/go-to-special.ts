/**
 * Go To Special — pure selection logic (Excel's F5 → Special).
 *
 * Given a worksheet's cell matrix and the active cell, compute the set of
 * ranges to select for a chosen criterion. Kept free of Univer/React so it's
 * unit-testable; the dialog feeds it the snapshot's `cellData` and applies the
 * result through `sheet.operation.set-selections`.
 *
 * Matched cells are coalesced into per-row contiguous column runs — a compact
 * multi-range selection that Univer renders as Excel does (disjoint blocks).
 */

export type GoToCriterion = 'constants' | 'formulas' | 'blanks' | 'currentRegion' | 'lastCell';

export interface SimpleRange {
  startRow: number;
  startColumn: number;
  endRow: number;
  endColumn: number;
}

/** Minimal cell shape we read from the snapshot's `cellData`. */
export interface GoToCell {
  v?: unknown;
  f?: string;
  p?: unknown;
  t?: number;
}

export type CellMatrix = Record<number, Record<number, GoToCell | undefined> | undefined>;

export interface GoToSpecialResult {
  /** Ranges to select; empty when no cells matched ("No cells were found."). */
  ranges: SimpleRange[];
}

const hasFormula = (c: GoToCell | undefined): boolean => typeof c?.f === 'string' && c.f !== '';

const hasValue = (c: GoToCell | undefined): boolean =>
  !!c && (c.p != null || (c.v != null && c.v !== ''));

/** A cell is a "constant" when it holds a literal value (or rich text) and is
 *  not driven by a formula. */
const isConstant = (c: GoToCell | undefined): boolean => !!c && !hasFormula(c) && hasValue(c);

/** A cell is "blank" when it carries no value, no rich text, and no formula. */
const isBlank = (c: GoToCell | undefined): boolean => !c || (!hasFormula(c) && !hasValue(c));

/** Used-range extent: the largest row/col index that holds any content. */
function usedBounds(cellData: CellMatrix): { maxRow: number; maxCol: number } {
  let maxRow = -1;
  let maxCol = -1;
  for (const rowKey of Object.keys(cellData)) {
    const row = cellData[Number(rowKey)];
    if (!row) continue;
    let rowHasContent = false;
    for (const colKey of Object.keys(row)) {
      const cell = row[Number(colKey)];
      if (isBlank(cell)) continue;
      rowHasContent = true;
      const col = Number(colKey);
      if (col > maxCol) maxCol = col;
    }
    if (rowHasContent) {
      const r = Number(rowKey);
      if (r > maxRow) maxRow = r;
    }
  }
  return { maxRow, maxCol };
}

/** Coalesce a set of matched cells into per-row contiguous column runs. */
function coalesce(matched: Array<{ row: number; col: number }>): SimpleRange[] {
  if (matched.length === 0) return [];
  // Group columns by row.
  const byRow = new Map<number, number[]>();
  for (const { row, col } of matched) {
    const cols = byRow.get(row);
    if (cols) cols.push(col);
    else byRow.set(row, [col]);
  }
  const ranges: SimpleRange[] = [];
  for (const [row, cols] of byRow) {
    cols.sort((a, b) => a - b);
    let start = cols[0];
    let prev = cols[0];
    for (let i = 1; i < cols.length; i++) {
      if (cols[i] === prev + 1) {
        prev = cols[i];
        continue;
      }
      ranges.push({ startRow: row, endRow: row, startColumn: start, endColumn: prev });
      start = cols[i];
      prev = cols[i];
    }
    ranges.push({ startRow: row, endRow: row, startColumn: start, endColumn: prev });
  }
  return ranges;
}

const getCell = (cellData: CellMatrix, row: number, col: number): GoToCell | undefined =>
  cellData[row]?.[col];

/** The contiguous block of non-empty cells around the active cell (Ctrl+* /
 *  Current region). 8-connected flood, then bounding box. */
function currentRegion(
  cellData: CellMatrix,
  active: { row: number; column: number },
  maxRow: number,
  maxCol: number,
): SimpleRange {
  // If the active cell is itself empty, Excel still selects the surrounding
  // block; we anchor on the active cell and expand into adjacent content.
  const seen = new Set<string>();
  const stack: Array<{ r: number; c: number }> = [{ r: active.row, c: active.column }];
  let minR = active.row;
  let maxR = active.row;
  let minC = active.column;
  let maxC = active.column;
  let found = false;
  while (stack.length) {
    const { r, c } = stack.pop()!;
    if (r < 0 || c < 0 || r > maxRow || c > maxCol) continue;
    const key = `${r}:${c}`;
    if (seen.has(key)) continue;
    seen.add(key);
    if (isBlank(getCell(cellData, r, c))) continue;
    found = true;
    minR = Math.min(minR, r);
    maxR = Math.max(maxR, r);
    minC = Math.min(minC, c);
    maxC = Math.max(maxC, c);
    for (let dr = -1; dr <= 1; dr++) {
      for (let dc = -1; dc <= 1; dc++) {
        if (dr === 0 && dc === 0) continue;
        stack.push({ r: r + dr, c: c + dc });
      }
    }
  }
  if (!found) {
    return {
      startRow: active.row,
      endRow: active.row,
      startColumn: active.column,
      endColumn: active.column,
    };
  }
  return { startRow: minR, endRow: maxR, startColumn: minC, endColumn: maxC };
}

export function computeGoToSpecial(
  cellData: CellMatrix,
  active: { row: number; column: number },
  criterion: GoToCriterion,
): GoToSpecialResult {
  const { maxRow, maxCol } = usedBounds(cellData);

  if (criterion === 'lastCell') {
    if (maxRow < 0 || maxCol < 0) return { ranges: [] };
    return {
      ranges: [{ startRow: maxRow, endRow: maxRow, startColumn: maxCol, endColumn: maxCol }],
    };
  }

  if (criterion === 'currentRegion') {
    if (maxRow < 0 || maxCol < 0) return { ranges: [] };
    return { ranges: [currentRegion(cellData, active, maxRow, maxCol)] };
  }

  if (maxRow < 0 || maxCol < 0) return { ranges: [] };

  const matched: Array<{ row: number; col: number }> = [];
  if (criterion === 'blanks') {
    // Blanks are only meaningful within the used range bounding box.
    for (let r = 0; r <= maxRow; r++) {
      for (let c = 0; c <= maxCol; c++) {
        if (isBlank(getCell(cellData, r, c))) matched.push({ row: r, col: c });
      }
    }
  } else {
    const pred = criterion === 'formulas' ? hasFormula : isConstant;
    for (const rowKey of Object.keys(cellData)) {
      const row = cellData[Number(rowKey)];
      if (!row) continue;
      for (const colKey of Object.keys(row)) {
        if (pred(row[Number(colKey)])) matched.push({ row: Number(rowKey), col: Number(colKey) });
      }
    }
  }

  return { ranges: coalesce(matched) };
}
