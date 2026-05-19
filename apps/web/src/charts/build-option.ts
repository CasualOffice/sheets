import type { FUniver } from '@univerjs/core/facade';
import type { ChartModel, ChartType } from './types';
import type { EChartsOption } from './echarts-init';

/**
 * Read cells from the chart's source range and turn them into an
 * ECharts option. Convention (mirrors Excel's default chart-from-
 * selection):
 *
 *   - Row 0 of the source range = header row → series names.
 *   - Column 0 = category axis labels (x-axis for column / area / line,
 *     y-axis for horizontal bar, dimension for pie).
 *   - Remaining cells = numeric values, one series per column.
 *
 * Non-numeric cells coerce to `null` so the chart shows a gap instead
 * of NaN. If the source range collapses to one row / one column we
 * fall back to a "No data" placeholder so the overlay still paints
 * rather than crashing.
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
    return { title: { text: 'No data', left: 'center', top: 'center' } };
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
  const seriesData: Array<Array<number | null>> = headers.map((_, sIdx) => {
    const data: Array<number | null> = [];
    for (let r = 1; r < rows; r++) {
      const v = ws.getRange(startRow + r, startColumn + 1 + sIdx).getValue();
      const n = typeof v === 'number' ? v : Number(v);
      data.push(Number.isFinite(n) ? n : null);
    }
    return data;
  });

  return buildOptionForType(model.type, headers, categories, seriesData, model.title);
}

function buildOptionForType(
  type: ChartType,
  headers: string[],
  categories: string[],
  rawSeries: Array<Array<number | null>>,
  title?: string,
): EChartsOption {
  if (type === 'pie' || type === 'doughnut') {
    // Excel's pie/doughnut uses the first value column only — match that
    // and label slices with the category column.
    const pieData = categories.map((label, i) => ({
      name: label,
      value: rawSeries[0]?.[i] ?? 0,
    }));
    return {
      title: title ? { text: title, left: 'center' } : undefined,
      tooltip: { trigger: 'item' },
      legend: { bottom: 0, type: 'scroll' },
      series: [
        {
          type: 'pie',
          radius: type === 'doughnut' ? ['40%', '70%'] : '60%',
          center: ['50%', title ? '52%' : '48%'],
          data: pieData,
          label: { formatter: '{b}' },
        },
      ],
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
    } as any;
  }

  if (type === 'scatter') {
    // Scatter expects [x, y] tuples. We pair the first value column
    // with each subsequent one — first column becomes the X axis,
    // remaining columns become Y series. Matches Excel's "Scatter
    // with only markers" picking column-0 as X.
    const xs = rawSeries[0] ?? [];
    const series = rawSeries.slice(1).map((ys, i) => ({
      name: headers[i + 1] ?? headers[0],
      type: 'scatter' as const,
      data: xs.map((x, idx) => [x, ys[idx]]).filter(([a, b]) => a != null && b != null),
    }));
    return {
      title: title ? { text: title, left: 'center' } : undefined,
      tooltip: { trigger: 'item' },
      legend: { bottom: 0, type: 'scroll' },
      grid: { left: 40, right: 16, top: title ? 40 : 16, bottom: 40 },
      xAxis: { type: 'value', name: headers[0] ?? '' },
      yAxis: { type: 'value' },
      series: series.length > 0 ? series : [{
        name: headers[0] ?? 'Series',
        type: 'scatter',
        data: xs.map((x, idx) => [idx, x]).filter(([, b]) => b != null),
      }],
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
    } as any;
  }

  // Bar / column / line / area share a category-axis + value-axis layout.
  const isHorizontalBar = type === 'bar' || type === 'bar-stacked' || type === 'bar-stacked-100';
  const is100 = type === 'column-stacked-100' || type === 'bar-stacked-100';
  const isStacked =
    type === 'column-stacked' ||
    type === 'column-stacked-100' ||
    type === 'bar-stacked' ||
    type === 'bar-stacked-100' ||
    type === 'line-stacked' ||
    type === 'area-stacked';
  const isLine = type === 'line' || type === 'line-stacked';
  const isArea = type === 'area' || type === 'area-stacked';
  const echartsType: 'bar' | 'line' = isLine || isArea ? 'line' : 'bar';

  // 100% stacked needs per-category sums for normalisation. Compute
  // them up front so we don't redo the work per series.
  const sumPerCat = is100
    ? categories.map((_, i) => {
        let s = 0;
        for (const ys of rawSeries) {
          const v = ys[i];
          if (typeof v === 'number') s += v;
        }
        return s === 0 ? 1 : s;
      })
    : null;

  const series = headers.map((name, sIdx) => {
    const raw = rawSeries[sIdx] ?? [];
    const data = is100 && sumPerCat
      ? raw.map((v, i) => (typeof v === 'number' ? (v / sumPerCat[i]) * 100 : null))
      : raw;
    return {
      name,
      type: echartsType,
      data,
      ...(isStacked ? { stack: 'all' as const } : {}),
      ...(isArea ? { areaStyle: {} } : {}),
      ...(isLine ? { smooth: false, symbol: 'circle' as const, symbolSize: 4 } : {}),
    };
  });

  const categoryAxis = { type: 'category' as const, data: categories };
  const valueAxis: Record<string, unknown> = { type: 'value' as const };
  if (is100) {
    valueAxis.max = 100;
    valueAxis.axisLabel = { formatter: '{value}%' };
  }

  return {
    title: title ? { text: title, left: 'center' } : undefined,
    tooltip: {
      trigger: 'axis',
      ...(is100 ? { valueFormatter: (v: unknown) => `${Math.round(Number(v))}%` } : {}),
    },
    legend: { bottom: 0, type: 'scroll' },
    grid: { left: 40, right: 16, top: title ? 40 : 16, bottom: 40 },
    xAxis: isHorizontalBar ? valueAxis : categoryAxis,
    yAxis: isHorizontalBar ? categoryAxis : valueAxis,
    series,
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
  } as any;
}
