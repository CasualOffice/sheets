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

import type { ChartModel } from './types';

/**
 * Excel auto-names charts as "Chart 1", "Chart 2", ..., picking the
 * lowest unused suffix even if intermediate charts have been deleted
 * (so deleting Chart 2 then inserting again gives you Chart 2, not
 * Chart 4). Match that.
 *
 * Operates on `ChartModel.title` — see the Chart Selection Pane for
 * inline rename.
 */
export function nextChartName(existing: ChartModel[]): string {
  const used = new Set<number>();
  for (const c of existing) {
    if (!c.title) continue;
    const m = /^Chart (\d+)$/.exec(c.title);
    if (m) used.add(Number(m[1]));
  }
  let n = 1;
  while (used.has(n)) n++;
  return `Chart ${n}`;
}
