import type { FUniver } from '@univerjs/core/facade';
import {
  PALETTES,
  mergeFormat,
  type ChartModel,
  type ChartType,
  type ResolvedChartFormat,
} from './types';
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
 *
 * Formatting (title visibility, legend position, axis titles,
 * gridlines, data labels, colour palette) is applied from
 * `mergeFormat(model)` — defaults match Excel's first-render.
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

  const format = mergeFormat(model);
  return buildOptionForType(model.type, headers, categories, seriesData, model.title, format);
}

function buildOptionForType(
  type: ChartType,
  headers: string[],
  categories: string[],
  rawSeries: Array<Array<number | null>>,
  title: string | undefined,
  format: ResolvedChartFormat,
): EChartsOption {
  const titleNode =
    title && format.showTitle ? { text: title, left: 'center' as const } : undefined;
  const colors = PALETTES[format.palette];
  const legendNode = legendOption(format);
  const showAxes = format.legend !== 'none';
  void showAxes;

  if (type === 'pie' || type === 'doughnut') {
    const pieData = categories.map((label, i) => ({
      name: label,
      value: rawSeries[0]?.[i] ?? 0,
    }));
    return {
      color: colors,
      title: titleNode,
      tooltip: { trigger: 'item' },
      legend: legendNode,
      series: [
        {
          type: 'pie',
          radius: type === 'doughnut' ? ['40%', '70%'] : '60%',
          center: ['50%', titleNode ? '52%' : '48%'],
          data: pieData,
          label: format.dataLabels
            ? { formatter: '{b}: {c}' }
            : { formatter: '{b}' },
        },
      ],
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
    } as any;
  }

  if (type === 'scatter') {
    const xs = rawSeries[0] ?? [];
    const series = rawSeries.slice(1).map((ys, i) => ({
      name: headers[i + 1] ?? headers[0],
      type: 'scatter' as const,
      data: xs.map((x, idx) => [x, ys[idx]]).filter(([a, b]) => a != null && b != null),
      label: dataLabelConfig(format),
    }));
    return {
      color: colors,
      title: titleNode,
      tooltip: { trigger: 'item' },
      legend: legendNode,
      grid: chartGrid(format, titleNode != null),
      xAxis: {
        type: 'value',
        name: format.xAxisTitle ?? headers[0] ?? '',
        nameLocation: 'middle',
        nameGap: 24,
        splitLine: { show: format.gridlines },
      },
      yAxis: {
        type: 'value',
        name: format.yAxisTitle ?? '',
        nameLocation: 'middle',
        nameGap: 36,
        splitLine: { show: format.gridlines },
      },
      series:
        series.length > 0
          ? series
          : [
              {
                name: headers[0] ?? 'Series',
                type: 'scatter' as const,
                data: xs.map((x, idx) => [idx, x]).filter(([, b]) => b != null),
                label: dataLabelConfig(format),
              },
            ],
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
    } as any;
  }

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
    // Linear-regression trendline overlay. We compute the regression
    // line endpoints in data space and emit a `markLine` from start to
    // end, dashed + slightly transparent so it doesn't compete with
    // the primary series.
    const trendlineMark = format.trendline
      ? buildTrendlineMark(data as Array<number | null>)
      : undefined;
    return {
      name,
      type: echartsType,
      data,
      ...(isStacked ? { stack: 'all' as const } : {}),
      ...(isArea ? { areaStyle: {} } : {}),
      ...(isLine ? { smooth: false, symbol: 'circle' as const, symbolSize: 4 } : {}),
      ...(trendlineMark ? { markLine: trendlineMark } : {}),
      label: dataLabelConfig(format, isHorizontalBar, isLine || isArea),
    };
  });

  const categoryAxis = {
    type: 'category' as const,
    data: categories,
    name: isHorizontalBar ? (format.yAxisTitle ?? '') : (format.xAxisTitle ?? ''),
    nameLocation: 'middle' as const,
    nameGap: 24,
  };
  const valueAxis: Record<string, unknown> = {
    type: 'value' as const,
    name: isHorizontalBar ? (format.xAxisTitle ?? '') : (format.yAxisTitle ?? ''),
    nameLocation: 'middle',
    nameGap: 36,
    splitLine: { show: format.gridlines },
  };
  if (is100) {
    valueAxis.max = 100;
    valueAxis.axisLabel = { formatter: '{value}%' };
  }

  return {
    color: colors,
    title: titleNode,
    tooltip: {
      trigger: 'axis',
      ...(is100 ? { valueFormatter: (v: unknown) => `${Math.round(Number(v))}%` } : {}),
    },
    legend: legendNode,
    grid: chartGrid(format, titleNode != null),
    xAxis: isHorizontalBar ? valueAxis : categoryAxis,
    yAxis: isHorizontalBar ? categoryAxis : valueAxis,
    series,
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
  } as any;
}

function legendOption(format: ResolvedChartFormat): Record<string, unknown> | undefined {
  if (format.legend === 'none') return undefined;
  const pos: Record<string, unknown> = { type: 'scroll' };
  switch (format.legend) {
    case 'top':
      pos.top = 0;
      break;
    case 'bottom':
      pos.bottom = 0;
      break;
    case 'left':
      pos.left = 0;
      pos.orient = 'vertical';
      break;
    case 'right':
      pos.right = 0;
      pos.orient = 'vertical';
      break;
  }
  return pos;
}

function chartGrid(
  format: ResolvedChartFormat,
  hasTitle: boolean,
): Record<string, unknown> {
  // Make room for legend / title / axis-name labels by padding the
  // plot area. Without this the value axis name gets clipped by the
  // legend at the bottom.
  const grid: Record<string, unknown> = {
    left: 56,
    right: 24,
    top: hasTitle ? 40 : 16,
    bottom: 56,
    containLabel: true,
  };
  switch (format.legend) {
    case 'top':
      grid.top = hasTitle ? 60 : 32;
      break;
    case 'left':
      grid.left = 96;
      break;
    case 'right':
      grid.right = 96;
      break;
    case 'none':
      grid.bottom = 32;
      break;
  }
  return grid;
}

function dataLabelConfig(
  format: ResolvedChartFormat,
  isHorizontalBar?: boolean,
  isLineOrArea?: boolean,
): Record<string, unknown> {
  if (!format.dataLabels) return { show: false };
  if (isLineOrArea) return { show: true, position: 'top' };
  if (isHorizontalBar) return { show: true, position: 'right' };
  return { show: true, position: 'top' };
}

/**
 * Linear regression trendline. Computes the best-fit line via simple
 * ordinary-least-squares on the series data points, then encodes it
 * as an ECharts `markLine` from (x_min, y_pred(x_min)) to (x_max,
 * y_pred(x_max)). Returns `undefined` if fewer than two valid
 * numeric points exist (a single point has no slope).
 */
function buildTrendlineMark(data: Array<number | null>): Record<string, unknown> | undefined {
  const points: Array<{ x: number; y: number }> = [];
  for (let i = 0; i < data.length; i += 1) {
    const v = data[i];
    if (typeof v === "number" && !Number.isNaN(v)) points.push({ x: i, y: v });
  }
  if (points.length < 2) return undefined;
  const n = points.length;
  let sumX = 0;
  let sumY = 0;
  let sumXY = 0;
  let sumXX = 0;
  for (const p of points) {
    sumX += p.x;
    sumY += p.y;
    sumXY += p.x * p.y;
    sumXX += p.x * p.x;
  }
  const denom = n * sumXX - sumX * sumX;
  if (denom === 0) return undefined;
  const slope = (n * sumXY - sumX * sumY) / denom;
  const intercept = (sumY - slope * sumX) / n;
  const xMin = points[0].x;
  const xMax = points[points.length - 1].x;
  const yAtMin = slope * xMin + intercept;
  const yAtMax = slope * xMax + intercept;
  return {
    silent: true,
    symbol: "none",
    lineStyle: { type: "dashed", width: 2, opacity: 0.75 },
    data: [[{ xAxis: xMin, yAxis: yAtMin }, { xAxis: xMax, yAxis: yAtMax }]],
  };
}
