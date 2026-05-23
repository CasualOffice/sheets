import type { FUniver } from '@univerjs/core/facade';
import type { PivotModel } from './types';
import type { PivotCell, SourceMatrix } from './compute';

/**
 * "Drill down" — given a click on a pivot result cell, return the
 * source records that contributed to it. Excel calls this "show
 * details" and dumps them into a fresh worksheet; we render them in a
 * popup instead (less disruption + no sheet sprawl).
 *
 * Cell-coordinate convention inside the pivot output rectangle:
 *
 *     [ row-field-name | Sum of … | Avg of … | ... ]     ← header (offsetRow = 0)
 *     [ row-key 1      |   …      |    …      | ... ]   ← offsetRow ≥ 1
 *     [ row-key 2      |   …      |    …      | ... ]
 *     [ Grand Total    |   …      |    …      | ... ]   ← offsetRow = lastRow
 *
 * Drilling on the header row is meaningless (returns null). Drilling
 * on the row-field column itself behaves the same as a Grand-Total
 * for that key. Drilling on Grand Total returns every filtered record.
 */

export type DrillDownResult = {
  /** Header labels matching the source columns. */
  headers: string[];
  /** Each contributing record, as a flat array of cell values. */
  rows: PivotCell[][];
  /** Friendly summary string, e.g. `Region = "North"` or
   *  `Grand Total · 12 rows`. Used as the dialog title. */
  summary: string;
};

/** Locate the pivot whose output rectangle contains `(row, col)` on
 *  the given sheet. Pivots without a recorded `lastOutputExtent` (old
 *  payloads from before P1 wrote it) are skipped — we can't bound
 *  their output without re-running compute. */
export function findPivotAtCell(
  pivots: PivotModel[],
  sheetId: string,
  row: number,
  col: number,
): PivotModel | null {
  for (const p of pivots) {
    if (p.targetSheetId !== sheetId) continue;
    const ext = p.lastOutputExtent;
    if (!ext) continue;
    const r0 = p.target.row;
    const c0 = p.target.column;
    if (row >= r0 && row < r0 + ext.rows && col >= c0 && col < c0 + ext.cols) {
      return p;
    }
  }
  return null;
}

/**
 * Compute the contributing rows for a click at absolute (row, col) on
 * the given pivot. Returns null if the click resolves to a non-
 * meaningful cell (the header row, or coordinates outside the pivot's
 * known extent).
 */
export function computeDrillDown(
  api: FUniver,
  pivot: PivotModel,
  row: number,
  col: number,
): DrillDownResult | null {
  // `col` is currently unused — drilling on any column within a row
  // returns the same set of contributing source records. Reserved
  // for a future per-value-field projection that only includes the
  // clicked value column.
  void col;
  const wb = api.getActiveWorkbook();
  if (!wb) return null;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const sheets = wb.getSheets() as any[];
  const sourceWs = sheets.find((s) => s.getSheetId?.() === pivot.sourceSheetId);
  if (!sourceWs) return null;

  const source = readSource(sourceWs, pivot.source);
  const offsetRow = row - pivot.target.row;
  const ext = pivot.lastOutputExtent;
  if (!ext) return null;
  if (offsetRow <= 0 || offsetRow >= ext.rows) return null; // header or out of bounds

  // Apply the same filters compute does, so drill-down rows match
  // what's visible in the pivot above.
  const filters = pivot.filters ?? [];
  const filtered = filters.length === 0
    ? source.records
    : source.records.filter((rec) => {
        for (const f of filters) {
          const allowed = new Set(f.allowedValues);
          const v = rec[f.column];
          const key = v == null ? '' : String(v);
          if (!allowed.has(key)) return false;
        }
        return true;
      });

  // Re-derive the row-key listing the same way compute.ts does so we
  // know which key the clicked row corresponds to. compute sorts the
  // keys ascending — match that order.
  const rowFieldCol = pivot.rows[0]?.column;
  const hasRowField = typeof rowFieldCol === 'number';
  const isGrandTotal = offsetRow === ext.rows - 1;

  if (isGrandTotal) {
    return {
      headers: source.headers,
      rows: filtered,
      summary: `Grand Total · ${filtered.length} rows`,
    };
  }

  if (!hasRowField) {
    // No row field + non-header / non-grand-total cell shouldn't
    // exist in compute's output, but if we get here, fall back to
    // returning every filtered record.
    return {
      headers: source.headers,
      rows: filtered,
      summary: `All rows · ${filtered.length} rows`,
    };
  }

  // Build the sorted key list and find the one at offsetRow - 1
  // (header occupies offsetRow = 0).
  const buckets = new Map<string, PivotCell[][]>();
  for (const rec of filtered) {
    const key = String(rec[rowFieldCol!] ?? '');
    let bucket = buckets.get(key);
    if (!bucket) {
      bucket = [];
      buckets.set(key, bucket);
    }
    bucket.push(rec);
  }
  const keys = [...buckets.keys()].sort((a, b) =>
    a.localeCompare(b, undefined, { numeric: true }),
  );
  const targetKey = keys[offsetRow - 1];
  if (targetKey == null) return null;
  const records = buckets.get(targetKey) ?? [];
  const fieldName = source.headers[rowFieldCol!] ?? 'value';
  return {
    headers: source.headers,
    rows: records,
    summary: `${fieldName} = "${targetKey || '(blank)'}" · ${records.length} rows`,
  };
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function readSource(ws: any, src: PivotModel['source']): SourceMatrix {
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
      if (v == null || v === '') row.push(null);
      else if (typeof v === 'number' || typeof v === 'string') {
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
    if (anyValue) records.push(row);
  }
  return { headers, records };
}
