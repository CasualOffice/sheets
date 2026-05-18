import { use } from 'echarts/core';
import { CanvasRenderer } from 'echarts/renderers';
import { BarChart, LineChart, PieChart, ScatterChart } from 'echarts/charts';
import {
  GridComponent,
  LegendComponent,
  TitleComponent,
  TooltipComponent,
  DatasetComponent,
} from 'echarts/components';

/**
 * Tree-shaken ECharts wiring. We pay only for the chart types + canvas
 * renderer we actually use. Pipeline Stage 4 already ships each lazy
 * group as its own chunk; this module is the chart-group entry point —
 * importing anything from `'echarts'` directly elsewhere would pull
 * the full library back in and undo the bundle savings.
 *
 * P0 ships Bar (the demo type). P3 layers Line / Pie / Scatter /
 * Area / Combo / Stacked variants on top — those use the same
 * Chart/Component registrations registered here, so future types
 * just import their chart constructor and add to `use([…])`.
 */
// `use` here is ECharts' renderer/component registry, not a React hook —
// silence rules-of-hooks for the module-level call.
// eslint-disable-next-line react-hooks/rules-of-hooks
use([
  CanvasRenderer,
  BarChart,
  LineChart,
  PieChart,
  ScatterChart,
  GridComponent,
  LegendComponent,
  TitleComponent,
  TooltipComponent,
  DatasetComponent,
]);

export { init } from 'echarts/core';
export type { EChartsType } from 'echarts/core';
export type { EChartsOption } from 'echarts';
