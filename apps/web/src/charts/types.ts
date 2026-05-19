/**
 * Chart model the in-memory store (`charts-context`) carries per
 * inserted chart. P0 is minimal — just enough to render a column
 * chart bound to a cell range. P1 adds the resource persistence
 * round-trip; P2 wires Univer's drawing model for move/resize; P3
 * expands `type` into the catalog of supported chart types.
 *
 * Position lives in cell coordinates rather than pixels — that's
 * how Excel anchors charts (top-left anchored to cell A, bottom-
 * right to cell B). We can compute pixels on the fly via
 * `range.getCellRect()` so the chart moves with the rows/columns
 * when they shift.
 */
/**
 * Chart subtypes, grouped by family. Matches the most-used subset of
 * Excel's "Insert Chart" catalog:
 *
 *   - Column (vertical bars): clustered / stacked / 100 %-stacked.
 *   - Bar (horizontal bars): clustered / stacked / 100 %-stacked.
 *   - Line: line / stacked line.
 *   - Area: area / stacked area.
 *   - Pie: pie / doughnut.
 *   - Scatter.
 *
 * The legacy `'bar'` literal that P0/P1 used for "vertical column
 * chart" is migrated to `'column'` on read (see `resources.ts`).
 */
export type ChartType =
  | 'column'
  | 'column-stacked'
  | 'column-stacked-100'
  | 'bar'
  | 'bar-stacked'
  | 'bar-stacked-100'
  | 'line'
  | 'line-stacked'
  | 'area'
  | 'area-stacked'
  | 'pie'
  | 'doughnut'
  | 'scatter';

/** Top-level family — what shows in the left column of the Insert dialog. */
export type ChartFamily = 'column' | 'bar' | 'line' | 'area' | 'pie' | 'scatter';

export const CHART_FAMILY_OF: Record<ChartType, ChartFamily> = {
  column: 'column',
  'column-stacked': 'column',
  'column-stacked-100': 'column',
  bar: 'bar',
  'bar-stacked': 'bar',
  'bar-stacked-100': 'bar',
  line: 'line',
  'line-stacked': 'line',
  area: 'area',
  'area-stacked': 'area',
  pie: 'pie',
  doughnut: 'pie',
  scatter: 'scatter',
};

/** Human-readable label shown in the panel + dialog. */
export const CHART_TYPE_LABEL: Record<ChartType, string> = {
  column: 'Clustered Column',
  'column-stacked': 'Stacked Column',
  'column-stacked-100': '100% Stacked Column',
  bar: 'Clustered Bar',
  'bar-stacked': 'Stacked Bar',
  'bar-stacked-100': '100% Stacked Bar',
  line: 'Line',
  'line-stacked': 'Stacked Line',
  area: 'Area',
  'area-stacked': 'Stacked Area',
  pie: 'Pie',
  doughnut: 'Doughnut',
  scatter: 'Scatter',
};

export type ChartModel = {
  id: string;
  /** Sheet the chart lives on. Matches `FWorksheet.getSheetId()`. */
  sheetId: string;
  /** Source data range — first row treated as header (column names),
   *  first column as category axis labels. */
  source: { startRow: number; endRow: number; startColumn: number; endColumn: number };
  /** Chart position, in 0-indexed cell coordinates. */
  pos: { startRow: number; endRow: number; startColumn: number; endColumn: number };
  type: ChartType;
  title?: string;
};

export function newChartId(): string {
  return `ch-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 6)}`;
}

/**
 * Plugin-resource name we use when stashing chart models in
 * `IWorkbookData.resources`. Mirrors `OUTLINE_RESOURCE_NAME` — survives
 * xlsx via the hidden `__casual_sheets_resources__` sheet and survives
 * collab via Univer's snapshot-load path.
 */
export const CHARTS_RESOURCE_NAME = '__casual_sheets_charts__';

/** Versioned envelope so a future schema change can be detected. */
export type ChartsResourceV1 = {
  v: 1;
  charts: ChartModel[];
};
