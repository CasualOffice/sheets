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

import type { FUniver } from '@univerjs/core/facade';
import { init } from './echarts-init';
import { buildEChartsOption } from './build-option';
import type { ChartModel } from './types';

/**
 * Render a chart model to a PNG ArrayBuffer using a detached ECharts
 * instance. Used by the xlsx exporter (Charts P5b) so the chart shows
 * up as a real image when the file is opened in Excel — the live
 * interactive ECharts canvas can't ship; an embedded image can.
 *
 * Constraints:
 *   - Must run on the main thread (ECharts needs a DOM container).
 *   - The container is created inside `document.body` with
 *     `visibility: hidden` + `pointer-events: none` so it never
 *     interferes with the actual UI even momentarily.
 *   - We choose pixel dimensions based on the chart's cell-anchored
 *     position so the embedded PNG roughly matches what the user
 *     sees in the app.
 *
 * Returns null when there's no source data (empty range etc.) so the
 * exporter can skip the chart cleanly rather than embed a 1×1 blank.
 */
export async function renderChartToPng(
  api: FUniver,
  model: ChartModel,
  pxWidth: number,
  pxHeight: number,
): Promise<ArrayBuffer | null> {
  const opt = buildEChartsOption(api, model);
  if (!opt) return null;

  const host = document.createElement('div');
  host.style.position = 'absolute';
  host.style.left = '-99999px';
  host.style.top = '0';
  host.style.width = `${pxWidth}px`;
  host.style.height = `${pxHeight}px`;
  host.style.visibility = 'hidden';
  host.style.pointerEvents = 'none';
  document.body.appendChild(host);

  try {
    const inst = init(host, undefined, { renderer: 'canvas' });
    // Apply the same option ChartOverlay uses so the snapshot matches
    // what the user sees in-app. notMerge: true to fully override any
    // ECharts defaults left over from the previous setOption.
    inst.setOption(opt, true);
    // ECharts renders on the next animation frame; force a synchronous
    // resize so the canvas is filled before we read it.
    inst.resize();
    const dataUrl = inst.getDataURL({
      type: 'png',
      pixelRatio: 2,
      backgroundColor: '#ffffff',
    });
    inst.dispose();
    return dataUrlToArrayBuffer(dataUrl);
  } finally {
    host.remove();
  }
}

function dataUrlToArrayBuffer(dataUrl: string): ArrayBuffer {
  const comma = dataUrl.indexOf(',');
  if (comma < 0) return new ArrayBuffer(0);
  const b64 = dataUrl.slice(comma + 1);
  const binary = atob(b64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes.buffer;
}

/**
 * Resolve the chart's pixel size from its cell-anchored position by
 * walking the sheet's column widths / row heights. Falls back to a
 * sensible default (480×320) when the snapshot doesn't have width/height
 * data for the spanned cells.
 */
export function pixelsForChart(
  api: FUniver,
  model: ChartModel,
): { width: number; height: number } {
  const wb = api.getActiveWorkbook();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const sheets = wb?.getSheets() as any[] | undefined;
  const ws = sheets?.find((s) => s.getSheetId?.() === model.sheetId);
  if (!ws) return { width: 480, height: 320 };

  let width = 0;
  for (let c = model.pos.startColumn; c <= model.pos.endColumn; c++) {
    const w = Number(ws.getColumnWidth?.(c) ?? 0);
    width += w > 0 ? w : 88;
  }
  let height = 0;
  for (let r = model.pos.startRow; r <= model.pos.endRow; r++) {
    const h = Number(ws.getRowHeight?.(r) ?? 0);
    height += h > 0 ? h : 24;
  }
  return {
    width: Math.max(120, Math.round(width)),
    height: Math.max(80, Math.round(height)),
  };
}
