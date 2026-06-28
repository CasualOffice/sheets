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
 * Pure, immutable model transforms for the PivotTable Fields pane.
 *
 * The pane mirrors Excel's "PivotTable Fields" panel: a source-field
 * list plus four drop zones — Filters / Columns / Rows / Values. Every
 * edit (assign a field to a zone, remove it, reorder, change a value's
 * aggregation) is expressed here as a `PivotModel → PivotModel` step so
 * the React panel never mutates the live model in place. (An earlier
 * bug elsewhere came from aliasing a live array; these clone instead.)
 *
 * Axis rules, matching Excel:
 *   - Rows / Columns / Filters hold a given source column at most once.
 *   - A field on an axis is mutually exclusive across Rows/Columns/Filters
 *     (moving it to one strips it from the others).
 *   - Values may hold the same column more than once (Sum + Count of the
 *     same field), so Values entries are addressed by index, not column.
 *   - Placing a field on an axis does NOT remove it from Values, and
 *     vice-versa (e.g. Region in Rows + Count of Region in Values).
 */

import type { DateGrouping, PivotAggregation, PivotModel, PivotValueField } from './types';

export type ZoneId = 'filters' | 'rows' | 'cols' | 'values';

export const ZONE_LABELS: Record<ZoneId, string> = {
  filters: 'Filters',
  cols: 'Columns',
  rows: 'Rows',
  values: 'Values',
};

/** Single-instance axes — a column appears at most once in each. */
const AXES: ZoneId[] = ['rows', 'cols', 'filters'];

/** Every source column placed somewhere, for the field-list checkmarks. */
export function placedColumns(model: PivotModel): Set<number> {
  const s = new Set<number>();
  for (const r of model.rows) s.add(r.column);
  for (const c of model.cols) s.add(c.column);
  for (const v of model.values) s.add(v.column);
  for (const f of model.filters ?? []) s.add(f.column);
  return s;
}

/** Shallow-clone the four zone arrays so callers can mutate the copies. */
function cloneZones(model: PivotModel): {
  rows: PivotModel['rows'];
  cols: PivotModel['cols'];
  values: PivotValueField[];
  filters: NonNullable<PivotModel['filters']>;
} {
  return {
    rows: model.rows.map((r) => ({ ...r })),
    cols: model.cols.map((c) => ({ ...c })),
    values: model.values.map((v) => ({ ...v })),
    filters: (model.filters ?? []).map((f) => ({ ...f, allowedValues: [...f.allowedValues] })),
  };
}

/** Drop a column from the single-instance axes (rows/cols/filters). */
function stripFromAxes(z: ReturnType<typeof cloneZones>, column: number): void {
  z.rows = z.rows.filter((r) => r.column !== column);
  z.cols = z.cols.filter((c) => c.column !== column);
  z.filters = z.filters.filter((f) => f.column !== column);
}

/**
 * Assign a source column to a zone.
 *
 * Rows/Columns/Filters: idempotent — moving the column to a new axis
 * strips it from the other two. Values: appends a new entry (so a field
 * can be aggregated more than once); `defaultAgg` lets the caller pick
 * Sum for numeric columns and Count for text.
 */
export function addFieldToZone(
  model: PivotModel,
  column: number,
  zone: ZoneId,
  opts?: { defaultAgg?: PivotAggregation; allowedValues?: string[] },
): PivotModel {
  const z = cloneZones(model);
  if (zone === 'values') {
    z.values.push({ column, agg: opts?.defaultAgg ?? 'sum', showAs: 'normal' });
    return { ...model, rows: z.rows, cols: z.cols, values: z.values, filters: z.filters };
  }
  // Axis zones are mutually exclusive — pull the column off every axis first.
  stripFromAxes(z, column);
  if (zone === 'rows') z.rows.push({ column });
  else if (zone === 'cols') z.cols.push({ column });
  else z.filters.push({ column, allowedValues: opts?.allowedValues ?? [] });
  return { ...model, rows: z.rows, cols: z.cols, values: z.values, filters: z.filters };
}

/** Remove the entry at `index` from a zone (Values by index; axes by index too). */
export function removeFieldFromZone(model: PivotModel, zone: ZoneId, index: number): PivotModel {
  const z = cloneZones(model);
  z[zone] = z[zone].filter((_, i) => i !== index) as never;
  return { ...model, rows: z.rows, cols: z.cols, values: z.values, filters: z.filters };
}

/** Reorder within a zone (drag a chip up/down, or the order buttons). */
export function moveWithinZone(
  model: PivotModel,
  zone: ZoneId,
  from: number,
  to: number,
): PivotModel {
  const z = cloneZones(model);
  const arr = z[zone] as unknown[];
  if (from < 0 || from >= arr.length || to < 0 || to >= arr.length || from === to) return model;
  const [moved] = arr.splice(from, 1);
  arr.splice(to, 0, moved);
  return { ...model, rows: z.rows, cols: z.cols, values: z.values, filters: z.filters };
}

/** Patch a Values entry (aggregation / Show-Values-As). */
export function updateValueField(
  model: PivotModel,
  index: number,
  patch: Partial<Pick<PivotValueField, 'agg' | 'showAs'>>,
): PivotModel {
  const values = model.values.map((v, i) => (i === index ? { ...v, ...patch } : v));
  return { ...model, values };
}

/** Set the date grouping on a Rows entry (Year/Quarter/Month/none). */
export function updateRowGrouping(
  model: PivotModel,
  index: number,
  grouping: DateGrouping,
): PivotModel {
  const rows = model.rows.map((r, i) =>
    i === index ? { ...r, grouping: grouping === 'none' ? undefined : grouping } : r,
  );
  return { ...model, rows };
}

/** True when the model still has at least one value field — the panel
 *  blocks removing the last one (a value-less pivot renders nothing). */
export function hasValues(model: PivotModel): boolean {
  return model.values.length > 0;
}

/** Which axis (if any) currently holds a column — drives the field-list
 *  badge ("R" / "C" / "▽"). Values is reported separately since a column
 *  can be on an axis and in Values at once. */
export function axisOf(model: PivotModel, column: number): Exclude<ZoneId, 'values'> | null {
  for (const ax of AXES) {
    const arr = ax === 'filters' ? (model.filters ?? []) : model[ax];
    if (arr.some((e) => e.column === column)) return ax as Exclude<ZoneId, 'values'>;
  }
  return null;
}
