import type { FUniver } from '@univerjs/core/facade';
import type { ChartModel } from './types';
import type { EChartsOption } from './echarts-init';

/**
 * Read cells from the chart's source range and turn them into an
 * ECharts option. Convention (mirrors Excel's default chart-from-
 * selection):
 *
 *   - Row 0 of the source range = header row → series names.
 *   - Column 0 = category axis labels (x-axis for column/bar charts,
 *     dimension for pie).
 *   - Remaining cells = numeric values, one series per column.
 *
 * Non-numeric cells are coerced to `null` so the chart shows a gap
 * instead of NaN. If the source range collapses to one row / one
 * column we still produce something sensible (single series).
 */
export function buildEChartsOption(api: FUniver, model: ChartModel): EChartsOption | null {
  const wb = api.getActiveWorkbook();
  if (!wb) return null;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const sheets = wb.getSheets() as any[];
  const ws = sheets.find((s) => s.getSheetId?.() === model.sheetId);
  if (!ws) return null;

  const { startRow, endRow, startColumn, endColumn } = model.source;
  const rows = endRow - startRow + 1;
  const cols = endColumn - startColumn + 1;
  if (rows < 2 || cols < 2) {
    // Need at least header row + 1 data row, and at least 1 label
    // column + 1 value column. Below that fall back to an empty
    // chart so the overlay still renders rather than crashing.
    return { title: { text: 'No data' } };
  }

  const headers: string[] = [];
  for (let c = 1; c < cols; c++) {
    const v = ws.getRange(startRow, startColumn + c).getValue();
    headers.push(v == null ? `Series ${c}` : String(v));
  }
  const categories: string[] = [];
  for (let r = 1; r < rows; r++) {
    const v = ws.getRange(startRow + r, startColumn).getValue();
    categories.push(v == null ? '' : String(v));
  }
  const series = headers.map((name, sIdx) => {
    const data: Array<number | null> = [];
    for (let r = 1; r < rows; r++) {
      const v = ws.getRange(startRow + r, startColumn + 1 + sIdx).getValue();
      const n = typeof v === 'number' ? v : Number(v);
      data.push(Number.isFinite(n) ? n : null);
    }
    return { name, type: model.type, data };
  });

  if (model.type === 'pie') {
    // Pie wants one series with {name, value} pairs. Use the first
    // value column.
    const pieData = categories.map((label, i) => ({
      name: label,
      value: series[0]?.data[i] ?? 0,
    }));
    return {
      title: model.title ? { text: model.title } : undefined,
      tooltip: { trigger: 'item' },
      legend: { bottom: 0 },
      series: [{ type: 'pie', radius: '60%', data: pieData }],
    };
  }

  return {
    title: model.title ? { text: model.title } : undefined,
    tooltip: { trigger: 'axis' },
    legend: { bottom: 0, type: 'scroll' },
    grid: { left: 40, right: 16, top: model.title ? 40 : 16, bottom: 40 },
    xAxis: { type: 'category', data: categories },
    yAxis: { type: 'value' },
    series,
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
  } as any;
}
