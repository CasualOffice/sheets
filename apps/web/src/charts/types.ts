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
export type ChartType = 'bar' | 'line' | 'pie' | 'scatter';

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
