import type { PivotAggregation, PivotModel } from './types';
import { PIVOT_AGG_LABELS } from './types';

/**
 * Pure compute step — turn raw source records into the laid-out cell
 * grid that gets written into the workbook. No FUniver dependency
 * here so it's trivially testable and so the same routine can run
 * server-side later (e.g. for an "export pivot to CSV" path).
 *
 * P0 scope:
 *   - Single row field (model.rows[0]). Multi-row keys land in P1.
 *   - No column field.
 *   - One or more value fields, each gets its own column in the
 *     output, with a Grand Total row at the bottom.
 *
 * Output cell layout (P0):
 *
 *     [ row-field-name | value-1 header | value-2 header | ... ]
 *     [ row-key 1      | agg(value-1)   | agg(value-2)   | ... ]
 *     [ row-key 2      | ...                                    ]
 *     [ Grand Total    | agg(all value-1) | agg(all value-2) | ]
 *
 * Cells are `string | number | null`. The apply step formats them
 * into Univer's IRange shape.
 */
export type PivotCell = string | number | null;
export type PivotGrid = PivotCell[][];

/**
 * Raw records read from the workbook source range. Row 0 is the
 * header row, rows 1.. are data. Cell values come in pre-coerced
 * by the caller (`apply.ts` reads them via `getValue()` which already
 * returns the right primitive type).
 */
export type SourceMatrix = {
  headers: string[];
  records: PivotCell[][];
};

export function computePivot(source: SourceMatrix, model: PivotModel): PivotGrid {
  if (model.values.length === 0) {
    // No value field — nothing to aggregate. Caller probably bailed
    // already, but return an empty grid so we don't crash if it didn't.
    return [];
  }

  const rowFieldCol = model.rows[0]?.column;
  const hasRowField = typeof rowFieldCol === 'number';

  // Bucket records by row key — when no row field is configured we
  // collapse everything into one anonymous bucket (Grand-Total-only).
  const buckets = new Map<string, PivotCell[][]>();
  for (const rec of source.records) {
    const key = hasRowField ? String(rec[rowFieldCol!] ?? '') : '';
    let bucket = buckets.get(key);
    if (!bucket) {
      bucket = [];
      buckets.set(key, bucket);
    }
    bucket.push(rec);
  }

  // Header row: row-field name (if any) + value-field headers like
  // "Sum of Sales" — matches Excel's auto-generated value-column
  // captions exactly.
  const header: PivotCell[] = [];
  if (hasRowField) {
    header.push(source.headers[rowFieldCol!] ?? '');
  }
  for (const v of model.values) {
    header.push(`${PIVOT_AGG_LABELS[v.agg]} of ${source.headers[v.column] ?? ''}`);
  }

  const grid: PivotGrid = [header];
  // Sort row keys ascending — Excel sorts alphabetically by default
  // when no explicit sort order is set.
  const keys = [...buckets.keys()].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
  for (const key of keys) {
    const rec = buckets.get(key)!;
    const row: PivotCell[] = [];
    if (hasRowField) row.push(key);
    for (const v of model.values) {
      row.push(aggregate(rec.map((r) => r[v.column]), v.agg));
    }
    grid.push(row);
  }

  // Grand total — even when there's no row field this provides the
  // single aggregated cell row.
  const total: PivotCell[] = [];
  if (hasRowField) total.push('Grand Total');
  for (const v of model.values) {
    total.push(aggregate(source.records.map((r) => r[v.column]), v.agg));
  }
  grid.push(total);

  return grid;
}

function aggregate(values: PivotCell[], agg: PivotAggregation): PivotCell {
  const nums: number[] = [];
  let nonNull = 0;
  for (const v of values) {
    if (v == null || v === '') continue;
    nonNull++;
    const n = typeof v === 'number' ? v : Number(v);
    if (Number.isFinite(n)) nums.push(n);
  }
  switch (agg) {
    case 'count':
      return nonNull;
    case 'sum':
      return nums.length === 0 ? 0 : nums.reduce((a, b) => a + b, 0);
    case 'average':
      return nums.length === 0 ? null : nums.reduce((a, b) => a + b, 0) / nums.length;
    case 'min':
      return nums.length === 0 ? null : Math.min(...nums);
    case 'max':
      return nums.length === 0 ? null : Math.max(...nums);
  }
}

/** Used by the panel to render a friendly auto-name like
 *  "Sum of Sales by Region" so a freshly inserted pivot has a label
 *  that explains itself. */
export function defaultPivotTitle(source: SourceMatrix, model: PivotModel): string {
  const value = model.values[0];
  if (!value) return 'PivotTable';
  const valuePart = `${PIVOT_AGG_LABELS[value.agg]} of ${source.headers[value.column] ?? 'value'}`;
  const rowField = model.rows[0];
  if (!rowField) return valuePart;
  return `${valuePart} by ${source.headers[rowField.column] ?? 'group'}`;
}
