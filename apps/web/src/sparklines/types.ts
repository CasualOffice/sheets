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

/**
 * Sparkline model — in-cell mini-charts. Each sparkline is anchored
 * to a single target cell (renders an SVG overlay sized to that
 * cell's bounding box) and reads its data from a range elsewhere in
 * the workbook.
 *
 * Three Excel-canonical types:
 *   - **line**: connects values with a small line; markers at min /
 *     max (auto-coloured).
 *   - **column**: short bars per value, positive in the series colour
 *     and negative tinted.
 *   - **win-loss**: tri-state bars (+ / − / 0) — used for binary
 *     up/down indicators.
 *
 * Persistence: sparklines round-trip through
 * `IWorkbookData.resources['__casual_sheets_sparklines__']`, same
 * pattern as charts and pivots. The resource sheet stays hidden in
 * xlsx output and is rehydrated on workbook reload.
 */

export type SparklineType = 'line' | 'column' | 'win-loss';

export type SparklineModel = {
  id: string;
  type: SparklineType;
  /** Workbook unit id this sparkline lives on. */
  unitId: string;
  /** Sheet id (subUnitId) of the anchor + source. v1 forces source +
   *  anchor to share a sheet; cross-sheet refs are a follow-up. */
  sheetId: string;
  /** Source data range (typically a single row or column). */
  source: { startRow: number; endRow: number; startColumn: number; endColumn: number };
  /** Anchor cell (row, col) — the sparkline renders inside this cell. */
  anchor: { row: number; col: number };
  /** Optional colour override; defaults to the accent green. */
  color?: string;
  /** Optional negative-bar colour for column / win-loss. */
  negativeColor?: string;
};

/** Plugin-resource name for stashing sparkline models in
 *  `IWorkbookData.resources`. Mirrors `PIVOTS_RESOURCE_NAME` — survives
 *  xlsx via the hidden `__casual_sheets_resources__` sheet. */
export const SPARKLINES_RESOURCE_NAME = '__casual_sheets_sparklines__';

export type SparklinesResourceV1 = {
  v: 1;
  sparklines: SparklineModel[];
};
