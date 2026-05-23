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
  /** Auto-generated "Chart N" name; the user can rename via the panel
   *  or the right-click menu. Doubles as the chart title shown above
   *  the plot when `format.showTitle` is true. */
  title?: string;
  /** Excel's "Format Chart Area" options — none required; absent
   *  values fall back to the defaults documented in `defaultFormat`. */
  format?: ChartFormat;
};

/** Equivalent of Excel's "Format Chart Area" pane in scope. */
export type ChartFormat = {
  /** Render the chart title above the plot. Defaults to `true` when
   *  the chart has a `title`, `false` otherwise. */
  showTitle?: boolean;
  /** Legend placement. `'none'` hides the legend entirely. */
  legend?: 'top' | 'right' | 'bottom' | 'left' | 'none';
  /** Optional axis titles. Ignored for pie / doughnut. */
  xAxisTitle?: string;
  yAxisTitle?: string;
  /** Show major gridlines on the value axis. */
  gridlines?: boolean;
  /** Show numeric labels on bars / lines / slices. */
  dataLabels?: boolean;
  /** Color palette applied to the series. Mirrors Excel's "Change
   *  Colors" picker — pick the first N from the chosen palette. */
  palette?: ChartPalette;
  /** Overlay a linear-regression trendline on every numeric series.
   *  Only meaningful for line / scatter / area / bar / column charts;
   *  ignored for pie. ECharts renders this via `markLine` from the
   *  computed regression endpoints. */
  trendline?: boolean;
};

export type ChartPalette = 'office' | 'mono' | 'vivid' | 'pastel';

/** Resolved (filled-in) format used by build-option. Apply over the
 *  user's partial via `mergeFormat`. */
export type ResolvedChartFormat = Required<Omit<ChartFormat, 'xAxisTitle' | 'yAxisTitle'>> & {
  xAxisTitle?: string;
  yAxisTitle?: string;
};

export function defaultFormat(model: Pick<ChartModel, 'title'>): ResolvedChartFormat {
  return {
    showTitle: Boolean(model.title),
    legend: 'bottom',
    gridlines: true,
    dataLabels: false,
    palette: 'office',
    trendline: false,
  };
}

export function mergeFormat(
  model: Pick<ChartModel, 'title' | 'format'>,
): ResolvedChartFormat {
  const base = defaultFormat(model);
  const f = model.format ?? {};
  return {
    showTitle: f.showTitle ?? base.showTitle,
    legend: f.legend ?? base.legend,
    gridlines: f.gridlines ?? base.gridlines,
    dataLabels: f.dataLabels ?? base.dataLabels,
    palette: f.palette ?? base.palette,
    trendline: f.trendline ?? base.trendline,
    xAxisTitle: f.xAxisTitle,
    yAxisTitle: f.yAxisTitle,
  };
}

/** Palette colours, applied in order to the series. Picked to match
 *  Excel's default colour sets so the look stays familiar. */
export const PALETTES: Record<ChartPalette, string[]> = {
  office: ['#5B9BD5', '#ED7D31', '#A5A5A5', '#FFC000', '#4472C4', '#70AD47', '#264478', '#9E480E'],
  mono:   ['#1F77B4', '#3F8FBC', '#5FA7C5', '#7FBFCD', '#9FD7D6', '#BFEFDE', '#5A8DAA', '#3D6E89'],
  vivid:  ['#E63946', '#F1A208', '#06A77D', '#005F73', '#9B5DE5', '#F15BB5', '#00BBF9', '#00F5D4'],
  pastel: ['#A3CEF1', '#FFD6A5', '#CAFFBF', '#FFADAD', '#BDB2FF', '#FDFFB6', '#FFC6FF', '#9BF6FF'],
};

export const PALETTE_LABELS: Record<ChartPalette, string> = {
  office: 'Office',
  mono:   'Monochromatic',
  vivid:  'Vivid',
  pastel: 'Pastel',
};

export const LEGEND_POSITIONS: { id: NonNullable<ChartFormat['legend']>; label: string }[] = [
  { id: 'bottom', label: 'Bottom' },
  { id: 'top',    label: 'Top' },
  { id: 'right',  label: 'Right' },
  { id: 'left',   label: 'Left' },
  { id: 'none',   label: 'None' },
];

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
