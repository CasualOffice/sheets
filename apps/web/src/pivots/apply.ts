import type { FUniver } from '@univerjs/core/facade';
import { computePivot, type SourceMatrix, type PivotGrid } from './compute';
import type { PivotModel } from './types';

/**
 * Read the source range, compute the pivot, then write the resulting
 * cell grid into the target sheet starting at `model.target`. The
 * write is one `setRangeValues` call so it lands as a single Univer
 * mutation — collab and undo both treat it atomically.
 *
 * The caller is responsible for clearing any previous pivot output
 * before calling apply again. P0 always writes a fresh pivot on
 * insert (no prior state to clear); P1's refresh path will clear
 * the previous output region first.
 */
export function applyPivot(api: FUniver, model: PivotModel): { rows: number; cols: number } | null {
  const wb = api.getActiveWorkbook();
  if (!wb) return null;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const sheets = wb.getSheets() as any[];
  const sourceWs = sheets.find((s) => s.getSheetId?.() === model.sourceSheetId);
  const targetWs = sheets.find((s) => s.getSheetId?.() === model.targetSheetId);
  if (!sourceWs || !targetWs) return null;

  const matrix = readSourceMatrix(sourceWs, model.source);
  if (matrix.records.length === 0) return null;

  const grid = computePivot(matrix, model);
  if (grid.length === 0) return null;

  writeGridToSheet(targetWs, model.target, grid);
  return { rows: grid.length, cols: grid[0]?.length ?? 0 };
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
