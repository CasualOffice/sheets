/**
 * Pivot table model. Excel's PivotCacheDefinition + PivotTableDefinition,
 * compressed to what we need for v0.1. The model is small and
 * JSON-serializable so it can live on `IWorkbookData.resources` the
 * same way chart models do.
 *
 * Conventions (same as charts):
 *
 *   - Source range row 0 = header row → field names.
 *   - Subsequent rows = records.
 *   - Field references are column indices within the source range
 *     (0-indexed). Names are looked up from the headers at render
 *     time so the model survives a column rename — there's no
 *     ambiguity until somebody changes the source range shape, at
 *     which point the field count may shift.
 */

export type PivotAggregation = 'sum' | 'count' | 'average' | 'min' | 'max' | 'distinctCount';

export const PIVOT_AGG_LABELS: Record<PivotAggregation, string> = {
  sum: 'Sum',
  count: 'Count',
  average: 'Average',
  min: 'Min',
  max: 'Max',
  distinctCount: 'Distinct Count',
};

/** Pivot field reference — column index within the source range. */
/** Date grouping for a (date) row field — buckets records by the derived
 *  period instead of the raw date. 'none' keys by the raw value. */
export type DateGrouping = 'none' | 'year' | 'quarter' | 'month';

export const PIVOT_DATE_GROUP_LABELS: Record<DateGrouping, string> = {
  none: 'No grouping',
  year: 'Years',
  quarter: 'Quarters',
  month: 'Months',
};

/** Pivot field reference — column index within the source range, plus an
 *  optional date grouping applied when bucketing rows. */
export type PivotFieldRef = { column: number; grouping?: DateGrouping };

/** Pivot value field — a column to aggregate + the aggregation. */
/** How a value field is displayed (Excel's "Show Values As"). 'normal' = the
 *  raw aggregate; 'pctOfGrandTotal' = each cell as a % of the field's overall
 *  grand total; 'pctOfColumnTotal' = each cell as a % of its column's total
 *  (in a cross-tab; identical to grand total in the row-only layout). */
export type PivotShowAs = 'normal' | 'pctOfGrandTotal' | 'pctOfColumnTotal';

export const PIVOT_SHOW_AS_LABELS: Record<PivotShowAs, string> = {
  normal: 'Normal',
  pctOfGrandTotal: '% of Grand Total',
  pctOfColumnTotal: '% of Column Total',
};

export type PivotValueField = {
  column: number;
  agg: PivotAggregation;
  /** Display transform; absent/`'normal'` shows the raw aggregate. */
  showAs?: PivotShowAs;
};

/** P1 — a filter field. Source records are kept only when the value
 *  in `column` is one of `allowedValues` (compared as strings). An
 *  empty `allowedValues` excludes everything; absent entry means no
 *  restriction. */
export type PivotFilter = {
  column: number;
  allowedValues: string[];
};

export type PivotModel = {
  id: string;
  /** Sheet the SOURCE data lives on. */
  sourceSheetId: string;
  /** Source data range — first row is headers, subsequent rows are
   *  records. */
  source: { startRow: number; endRow: number; startColumn: number; endColumn: number };
  /** Sheet the pivot OUTPUT lives on. Often the same as source, but
   *  Excel lets you target a different sheet. */
  targetSheetId: string;
  /** Top-left cell of the pivot output (0-indexed row + col). */
  target: { row: number; column: number };
  /** Row fields, applied left-to-right. P0 ships single-field; the
   *  array shape makes multi-field a P1 follow-up without a schema
   *  bump. */
  rows: PivotFieldRef[];
  /** Column fields. Empty in P0 — every value column collapses into
   *  one column. */
  cols: PivotFieldRef[];
  /** Value fields. Each one produces its own column in the output. */
  values: PivotValueField[];
  /** P1 — filter fields. Optional for backwards-compat; an absent or
   *  empty array means no filtering. */
  filters?: PivotFilter[];
  /** Last written output extent so `Refresh` can clear the previous
   *  output rectangle before writing the new one. Absent on freshly-
   *  loaded pivots from pre-P1 workbooks — refresh in that case just
   *  overwrites the new extent and leaves any residual rows. */
  lastOutputExtent?: { rows: number; cols: number };
  /** Display name. Auto-generated "PivotTable N" on insert; renameable
   *  via the Pivots panel. */
  title?: string;
};

export function newPivotId(): string {
  return `pt-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 6)}`;
}

/** Plugin-resource name we use when stashing pivot models in
 *  `IWorkbookData.resources`. Mirrors `CHARTS_RESOURCE_NAME` — survives
 *  xlsx via the hidden `__casual_sheets_resources__` sheet. */
export const PIVOTS_RESOURCE_NAME = '__casual_sheets_pivots__';

export type PivotsResourceV1 = {
  v: 1;
  pivots: PivotModel[];
};
