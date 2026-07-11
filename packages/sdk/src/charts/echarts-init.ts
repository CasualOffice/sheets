/**
 * Copyright 2026 Casual Office
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

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
