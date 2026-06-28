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

import type { IWorkbookData } from '@univerjs/core';
import {
  CHARTS_RESOURCE_NAME,
  type ChartModel,
  type ChartsResourceV1,
  type ChartType,
} from './types';

/**
 * Round-trip helpers for the charts plugin resource. The resource lives on
 * `IWorkbookData.resources` and is carried through xlsx via the hidden
 * `__casual_sheets_resources__` sheet we already use for outline groups,
 * and through collab via Univer's snapshot-load path.
 */

const VALID_TYPES: ChartType[] = [
  'column',
  'column-stacked',
  'column-stacked-100',
  'bar',
  'bar-stacked',
  'bar-stacked-100',
  'line',
  'line-stacked',
  'area',
  'area-stacked',
  'pie',
  'doughnut',
  'scatter',
];

/**
 * The pre-P3 store used `'bar'` for what we now (correctly) call
 * `'column'` — vertical bars. Migrate on read so existing workbooks
 * keep rendering after the rename.
 */
function migrateType(raw: unknown): ChartType | null {
  if (typeof raw !== 'string') return null;
  // Existing 'bar' string is ambiguous between old "vertical column"
  // and new "horizontal bar". P0/P1/P2 saved 'bar' meaning column;
  // we have no horizontal-bar charts in the wild yet, so 'bar' from
  // before P3 means column.
  if (raw === 'bar') return 'column';
  return VALID_TYPES.includes(raw as ChartType) ? (raw as ChartType) : null;
}

function isValidChart(c: unknown): c is ChartModel {
  if (!c || typeof c !== 'object') return false;
  const r = c as Record<string, unknown>;
  if (typeof r.id !== 'string' || typeof r.sheetId !== 'string') return false;
  const migrated = migrateType(r.type);
  if (!migrated) return false;
  r.type = migrated; // mutate so the rest of the app sees the new value
  const src = r.source as Record<string, unknown> | undefined;
  const pos = r.pos as Record<string, unknown> | undefined;
  if (!src || !pos) return false;
  for (const k of ['startRow', 'endRow', 'startColumn', 'endColumn'] as const) {
    if (typeof src[k] !== 'number' || typeof pos[k] !== 'number') return false;
  }
  // `format` is optional and freely shaped — defer validation to
  // `mergeFormat`, which fills missing fields with defaults. Anything
  // we don't recognise is ignored at render time.
  if (r.format != null && typeof r.format !== 'object') return false;
  return true;
}

/** Read chart models out of a snapshot. Tolerant of older / missing payloads. */
export function readChartsFromSnapshot(data: IWorkbookData | undefined): ChartModel[] {
  if (!data?.resources?.length) return [];
  const entry = data.resources.find((r) => r.name === CHARTS_RESOURCE_NAME);
  if (!entry?.data) return [];
  try {
    const parsed = JSON.parse(entry.data) as Partial<ChartsResourceV1>;
    if (parsed?.v !== 1 || !Array.isArray(parsed.charts)) return [];
    return parsed.charts.filter(isValidChart);
  } catch {
    /* corrupt payload — drop silently, the workbook still opens fine */
    return [];
  }
}

/** Merge chart models INTO `data.resources` for export. Mutates in place. */
export function writeChartsIntoSnapshot(
  data: IWorkbookData,
  charts: ChartModel[],
): void {
  const existing = data.resources ?? [];
  const filtered = existing.filter((r) => r.name !== CHARTS_RESOURCE_NAME);
  if (charts.length === 0) {
    data.resources = filtered;
    return;
  }
  const payload: ChartsResourceV1 = { v: 1, charts };
  data.resources = [...filtered, { name: CHARTS_RESOURCE_NAME, data: JSON.stringify(payload) }];
}
