/**
 * Copyright 2026 Casual Office
 *
 * Licensed under the Apache License, Version 2.0 (the "License").
 */

/**
 * PivotTable Fields side panel — Excel's "PivotTable Fields" task pane, ported
 * from the standalone app's PivotFieldsPanel into the SDK chrome so embedders
 * (dochub) get it natively.
 *
 * Lists the source fields of a pivot and its four drop zones — Filters /
 * Columns / Rows / Values — and lets the user assign fields (drag, or the "+"
 * zone menu), remove/reorder chips, edit a value's aggregation + Show-Values-As,
 * a row field's date grouping, and a report filter's per-value checklist. The
 * pivot config is persisted on the workbook so it round-trips through xlsx and
 * collab.
 *
 * DECOUPLING from the app:
 *   - `useUniverAPI()` → the injected `api` prop; `useUI().togglePivotPanel` →
 *     the `onClose` prop.
 *   - The app reads/writes pivot models through its `PivotsProvider` context
 *     (apps/web/src/pivots/pivots-context.tsx), which does NOT exist in the SDK.
 *     Instead — mirroring the SDK's InsertSparklineDialog — we read the pivot
 *     models straight off the live workbook snapshot's
 *     `resources['__casual_sheets_pivots__']` via `api.getContent()` and persist
 *     edits back with `api.setContent()`. Same resource envelope the app's
 *     pivots feature owns; it round-trips through xlsx + collab identically.
 *   - The pure `fields-model` transforms + label tables + source reader are
 *     INLINED here (structurally identical to apps/web/src/pivots/*) so the SDK
 *     stays free of the app package.
 *
 * COMPUTE GAP (honest limitation): the app's pivot *engine* — `computePivot` /
 * `applyPivot` (apps/web/src/pivots/{compute,apply}.ts) — is app-only code, not
 * a Univer plugin, and porting it is out of scope for this PR. This panel fully
 * edits + persists the pivot's field configuration (definition), but it does NOT
 * re-render the output cell grid: the laid-out pivot rectangle on the sheet only
 * refreshes once the app engine re-applies the model. So field assignment,
 * drag-and-drop, agg/showAs/grouping/filters all work and are saved, but the
 * visible output lags the config until the engine port lands (a later PR).
 */

import { useEffect, useMemo, useState, type CSSProperties, type ReactNode } from 'react';
import type { IWorkbookData } from '@univerjs/core';

import type { CasualSheetsAPI } from '../sheets/api';
import type { PanelComponentProps } from './extensions';
import { Icon } from './Icon';

/* ------------------------------------------------------------------ *
 * Pivot model — mirrors apps/web/src/pivots/types.ts (duplicated so the
 * SDK doesn't depend on the app package).
 * ------------------------------------------------------------------ */

type PivotAggregation = 'sum' | 'count' | 'average' | 'min' | 'max' | 'distinctCount';
const PIVOT_AGG_LABELS: Record<PivotAggregation, string> = {
  sum: 'Sum',
  count: 'Count',
  average: 'Average',
  min: 'Min',
  max: 'Max',
  distinctCount: 'Distinct Count',
};

type DateGrouping = 'none' | 'year' | 'quarter' | 'month';
const PIVOT_DATE_GROUP_LABELS: Record<DateGrouping, string> = {
  none: 'No grouping',
  year: 'Years',
  quarter: 'Quarters',
  month: 'Months',
};

type PivotShowAs = 'normal' | 'pctOfGrandTotal' | 'pctOfColumnTotal' | 'pctOfRowTotal';
const PIVOT_SHOW_AS_LABELS: Record<PivotShowAs, string> = {
  normal: 'Normal',
  pctOfGrandTotal: '% of Grand Total',
  pctOfColumnTotal: '% of Column Total',
  pctOfRowTotal: '% of Row Total',
};

type PivotFieldRef = { column: number; grouping?: DateGrouping };
type PivotValueField = { column: number; agg: PivotAggregation; showAs?: PivotShowAs };
type PivotFilter = { column: number; allowedValues: string[] };

interface PivotModel {
  id: string;
  sourceSheetId: string;
  source: { startRow: number; endRow: number; startColumn: number; endColumn: number };
  targetSheetId: string;
  target: { row: number; column: number };
  rows: PivotFieldRef[];
  cols: PivotFieldRef[];
  values: PivotValueField[];
  filters?: PivotFilter[];
  lastOutputExtent?: { rows: number; cols: number };
  title?: string;
}

/** Resource key + envelope — mirrors `PIVOTS_RESOURCE_NAME` /
 *  `PivotsResourceV1` in apps/web/src/pivots/types.ts. */
const PIVOTS_RESOURCE_NAME = '__casual_sheets_pivots__';
interface PivotsResourceV1 {
  v: 1;
  pivots: PivotModel[];
}

/* ------------------------------------------------------------------ *
 * Pure model transforms — mirrors apps/web/src/pivots/fields-model.ts.
 * ------------------------------------------------------------------ */

type ZoneId = 'filters' | 'rows' | 'cols' | 'values';

const ZONE_LABELS: Record<ZoneId, string> = {
  filters: 'Filters',
  cols: 'Columns',
  rows: 'Rows',
  values: 'Values',
};

const AXES: Array<Exclude<ZoneId, 'values'>> = ['rows', 'cols', 'filters'];

type DragPayload = { from: 'list'; column: number } | { from: 'zone'; zone: ZoneId; index: number };

function placedColumns(model: PivotModel): Set<number> {
  const s = new Set<number>();
  for (const r of model.rows) s.add(r.column);
  for (const c of model.cols) s.add(c.column);
  for (const v of model.values) s.add(v.column);
  for (const f of model.filters ?? []) s.add(f.column);
  return s;
}

function cloneZones(model: PivotModel) {
  return {
    rows: model.rows.map((r) => ({ ...r })),
    cols: model.cols.map((c) => ({ ...c })),
    values: model.values.map((v) => ({ ...v })),
    filters: (model.filters ?? []).map((f) => ({ ...f, allowedValues: [...f.allowedValues] })),
  };
}

function stripFromAxes(z: ReturnType<typeof cloneZones>, column: number): void {
  z.rows = z.rows.filter((r) => r.column !== column);
  z.cols = z.cols.filter((c) => c.column !== column);
  z.filters = z.filters.filter((f) => f.column !== column);
}

function addFieldToZone(
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
  stripFromAxes(z, column);
  if (zone === 'rows') z.rows.push({ column });
  else if (zone === 'cols') z.cols.push({ column });
  else z.filters.push({ column, allowedValues: opts?.allowedValues ?? [] });
  return { ...model, rows: z.rows, cols: z.cols, values: z.values, filters: z.filters };
}

function removeFieldFromZone(model: PivotModel, zone: ZoneId, index: number): PivotModel {
  const z = cloneZones(model);
  z[zone] = (z[zone] as unknown[]).filter((_, i) => i !== index) as never;
  return { ...model, rows: z.rows, cols: z.cols, values: z.values, filters: z.filters };
}

function moveWithinZone(model: PivotModel, zone: ZoneId, from: number, to: number): PivotModel {
  const z = cloneZones(model);
  const arr = z[zone] as unknown[];
  if (from < 0 || from >= arr.length || to < 0 || to >= arr.length || from === to) return model;
  const [moved] = arr.splice(from, 1);
  arr.splice(to, 0, moved);
  return { ...model, rows: z.rows, cols: z.cols, values: z.values, filters: z.filters };
}

function updateValueField(
  model: PivotModel,
  index: number,
  patch: Partial<Pick<PivotValueField, 'agg' | 'showAs'>>,
): PivotModel {
  const values = model.values.map((v, i) => (i === index ? { ...v, ...patch } : v));
  return { ...model, values };
}

function updateRowGrouping(model: PivotModel, index: number, grouping: DateGrouping): PivotModel {
  const rows = model.rows.map((r, i) =>
    i === index ? { ...r, grouping: grouping === 'none' ? undefined : grouping } : r,
  );
  return { ...model, rows };
}

function hasValues(model: PivotModel): boolean {
  return model.values.length > 0;
}

function axisOf(model: PivotModel, column: number): Exclude<ZoneId, 'values'> | null {
  for (const ax of AXES) {
    const arr = ax === 'filters' ? (model.filters ?? []) : model[ax];
    if (arr.some((e) => e.column === column)) return ax;
  }
  return null;
}

function toggleFilterValue(
  model: PivotModel,
  filterIndex: number,
  value: string,
  checked: boolean,
  allValues: string[],
): PivotModel {
  const filters = (model.filters ?? []).map((f) => ({ ...f, allowedValues: [...f.allowedValues] }));
  const f = filters[filterIndex];
  if (!f) return model;
  const current = new Set(f.allowedValues.length ? f.allowedValues : allValues);
  if (checked) current.add(value);
  else current.delete(value);
  f.allowedValues = allValues.filter((v) => current.has(v));
  return { ...model, filters };
}

function setFilterValues(model: PivotModel, filterIndex: number, values: string[]): PivotModel {
  const filters = (model.filters ?? []).map((f) => ({ ...f, allowedValues: [...f.allowedValues] }));
  if (!filters[filterIndex]) return model;
  filters[filterIndex].allowedValues = [...values];
  return { ...model, filters };
}

function filterAllowedCount(model: PivotModel, filterIndex: number, allCount: number): number {
  const f = (model.filters ?? [])[filterIndex];
  if (!f) return allCount;
  return f.allowedValues.length ? f.allowedValues.length : allCount;
}

function columnInZone(model: PivotModel, zone: ZoneId, index: number): number | null {
  const arr =
    zone === 'values' ? model.values : zone === 'filters' ? (model.filters ?? []) : model[zone];
  return arr[index]?.column ?? null;
}

function applyDrop(
  model: PivotModel,
  payload: DragPayload,
  targetZone: ZoneId,
  optsFor: (column: number) => { defaultAgg?: PivotAggregation; allowedValues?: string[] },
): PivotModel {
  if (payload.from === 'list') {
    return addFieldToZone(model, payload.column, targetZone, optsFor(payload.column));
  }
  if (payload.zone === targetZone) return model;
  const column = columnInZone(model, payload.zone, payload.index);
  if (column == null) return model;
  if (payload.zone === 'values' && model.values.length <= 1) return model;
  const removed = removeFieldFromZone(model, payload.zone, payload.index);
  return addFieldToZone(removed, column, targetZone, optsFor(column));
}

/* ------------------------------------------------------------------ *
 * Snapshot resource read/write — the SDK-native persistence path
 * (mirrors InsertSparklineDialog): read models off `api.getContent()`,
 * write back through `api.setContent()`.
 * ------------------------------------------------------------------ */

function readPivots(api: CasualSheetsAPI): PivotModel[] {
  const data = api.getContent();
  const entry = data?.resources?.find((r) => r.name === PIVOTS_RESOURCE_NAME);
  if (!entry?.data) return [];
  try {
    const parsed = JSON.parse(entry.data) as Partial<PivotsResourceV1>;
    if (parsed?.v !== 1 || !Array.isArray(parsed.pivots)) return [];
    return parsed.pivots;
  } catch {
    return [];
  }
}

/** Persist an edited pivot model back onto the workbook snapshot. Merges over
 *  the existing entry and re-mounts via `setContent` — the same path the
 *  sparkline dialog uses. */
function persistPivot(api: CasualSheetsAPI, next: PivotModel): void {
  const data = api.getContent();
  if (!data) return;
  const resources = data.resources ? [...data.resources] : [];
  const idx = resources.findIndex((r) => r.name === PIVOTS_RESOURCE_NAME);
  let pivots: PivotModel[] = [];
  if (idx >= 0 && resources[idx]?.data) {
    try {
      const parsed = JSON.parse(resources[idx].data) as Partial<PivotsResourceV1>;
      if (parsed?.v === 1 && Array.isArray(parsed.pivots)) pivots = parsed.pivots;
    } catch {
      pivots = [];
    }
  }
  const merged = pivots.some((p) => p.id === next.id)
    ? pivots.map((p) => (p.id === next.id ? next : p))
    : [...pivots, next];
  const entry = { name: PIVOTS_RESOURCE_NAME, data: JSON.stringify({ v: 1, pivots: merged }) };
  if (idx >= 0) resources[idx] = entry;
  else resources.push(entry);
  const nextData: IWorkbookData = { ...data, resources };
  api.setContent(nextData);
}

/* ------------------------------------------------------------------ *
 * Source reader — headers + distinct values off the FUniver facade.
 * Mirrors readSource() in apps/web/src/pivots/PivotFieldsPanel.tsx.
 * ------------------------------------------------------------------ */

interface SourceView {
  headers: string[];
  distinct: (col: number) => string[];
  isNumeric: (col: number) => boolean;
}

/* eslint-disable @typescript-eslint/no-explicit-any */
function readSource(api: CasualSheetsAPI, model: PivotModel | null): SourceView {
  const empty: SourceView = { headers: [], distinct: () => [], isNumeric: () => false };
  if (!model) return empty;
  const wb = api.univer.getActiveWorkbook?.();
  if (!wb) return empty;
  const ws = (wb.getSheets() as any[]).find((s) => s.getSheetId?.() === model.sourceSheetId);
  if (!ws) return empty;
  const { startRow, endRow, startColumn, endColumn } = model.source;
  const headers: string[] = [];
  for (let c = startColumn; c <= endColumn; c++) {
    const v = ws.getRange(startRow, c).getValue();
    headers.push(v == null || v === '' ? `Column ${c - startColumn + 1}` : String(v));
  }
  const colAt = (idx: number) => startColumn + idx;
  return {
    headers,
    distinct: (col) => {
      const seen = new Set<string>();
      for (let r = startRow + 1; r <= endRow; r++) {
        const v = ws.getRange(r, colAt(col)).getValue();
        seen.add(v == null ? '' : String(v));
      }
      return [...seen].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
    },
    isNumeric: (col) => {
      for (let r = startRow + 1; r <= endRow; r++) {
        const v = ws.getRange(r, colAt(col)).getValue();
        if (v == null || v === '') continue;
        return typeof v === 'number';
      }
      return false;
    },
  };
}
/* eslint-enable @typescript-eslint/no-explicit-any */

/* ------------------------------------------------------------------ *
 * Styles — inline, --cs-chrome-* CSS vars, no external classes.
 * ------------------------------------------------------------------ */

const border = '1px solid var(--cs-chrome-border, #edeff3)';

const headerStyle: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 8,
  padding: '10px 12px',
  borderBottom: border,
};

const iconBtn: CSSProperties = {
  border: 'none',
  background: 'transparent',
  cursor: 'pointer',
  color: 'inherit',
  display: 'inline-flex',
  alignItems: 'center',
  padding: 2,
};

const sectionTitle: CSSProperties = {
  margin: '0 0 6px',
  fontSize: 11,
  fontWeight: 700,
  textTransform: 'uppercase',
  letterSpacing: 0.4,
  opacity: 0.6,
};

const fieldRow: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 6,
  padding: '5px 6px',
  borderRadius: 6,
  cursor: 'grab',
  userSelect: 'none',
};

const chipStyle: CSSProperties = {
  border,
  borderRadius: 6,
  padding: '6px 8px',
  background: 'var(--cs-chrome-surface, #fbfbfd)',
  display: 'flex',
  flexDirection: 'column',
  gap: 6,
};

const selectStyle: CSSProperties = {
  font: 'inherit',
  fontSize: 12,
  padding: '2px 4px',
  border,
  borderRadius: 4,
  background: 'transparent',
  color: 'inherit',
  flex: 1,
};

/* ------------------------------------------------------------------ *
 * Panel component.
 * ------------------------------------------------------------------ */

export function PivotFieldsPanel({ api, onClose }: PanelComponentProps) {
  const [pivots, setPivots] = useState<PivotModel[]>(() => readPivots(api));
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [menuFor, setMenuFor] = useState<number | null>(null);
  const [expandedFilter, setExpandedFilter] = useState<number | null>(null);

  // Re-read pivots whenever the workbook content settles (edit / import /
  // setContent / collab). Keeps the panel live without the app pivot store.
  useEffect(() => {
    const reload = () => setPivots(readPivots(api));
    reload();
    const offChange = api.on('change', reload);
    return () => offChange();
  }, [api]);

  // Keep a valid selection: default to the most-recent pivot; recover if the
  // selected one was deleted.
  useEffect(() => {
    if (pivots.length === 0) {
      if (selectedId !== null) setSelectedId(null);
      return;
    }
    if (!selectedId || !pivots.some((p) => p.id === selectedId)) {
      setSelectedId(pivots[pivots.length - 1].id);
    }
  }, [pivots, selectedId]);

  const model = useMemo(
    () => pivots.find((p) => p.id === selectedId) ?? null,
    [pivots, selectedId],
  );
  const source = useMemo(() => readSource(api, model), [api, model]);

  const placed = model ? placedColumns(model) : new Set<number>();

  // Optimistic local update + persist. We update `pivots` immediately so the UI
  // is responsive, then write through to the snapshot so it round-trips.
  const commit = (next: PivotModel) => {
    setPivots((prev) => prev.map((p) => (p.id === next.id ? next : p)));
    persistPivot(api, next);
  };

  const optsFor = (col: number) => ({
    defaultAgg: source.isNumeric(col) ? ('sum' as const) : ('count' as const),
    allowedValues: source.distinct(col),
  });

  const onZoneDrop = (zone: ZoneId, raw: string) => {
    if (!model || !raw) return;
    let payload: DragPayload;
    try {
      payload = JSON.parse(raw) as DragPayload;
    } catch {
      return;
    }
    commit(applyDrop(model, payload, zone, optsFor));
  };

  const labelFor = (col: number) => source.headers[col] ?? `Column ${col + 1}`;

  return (
    <div
      data-testid="cs-pivot-fields-panel"
      style={{ height: '100%', display: 'flex', flexDirection: 'column' }}
    >
      <header style={headerStyle}>
        <Icon name="pivot_table_chart" size={18} />
        <span style={{ fontWeight: 600, flex: 1 }}>PivotTable Fields</span>
        <button
          type="button"
          aria-label="Close PivotTable Fields panel"
          onClick={onClose}
          style={iconBtn}
        >
          <Icon name="close" size={18} />
        </button>
      </header>

      <div style={{ flex: 1, overflow: 'auto', padding: 12 }}>
        {!model ? (
          <div
            data-testid="cs-pivot-fields-empty"
            style={{ textAlign: 'center', opacity: 0.75, padding: '24px 8px' }}
          >
            <Icon name="pivot_table_chart" size={40} style={{ opacity: 0.4 }} />
            <div style={{ fontWeight: 600, marginTop: 8 }}>No PivotTable selected</div>
            <div style={{ fontSize: 13, marginTop: 4 }}>
              Insert a PivotTable, then configure its fields here.
            </div>
          </div>
        ) : (
          <>
            {pivots.length > 1 && (
              <label style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 12 }}>
                <span style={{ fontSize: 12, opacity: 0.7 }}>PivotTable</span>
                <select
                  data-testid="cs-pivot-fields-picker"
                  value={selectedId ?? ''}
                  onChange={(e) => setSelectedId(e.target.value)}
                  style={{ ...selectStyle, flex: 1 }}
                >
                  {pivots.map((p, i) => (
                    <option key={p.id} value={p.id}>
                      {p.title ?? `PivotTable ${i + 1}`}
                    </option>
                  ))}
                </select>
              </label>
            )}

            {/* Compute-gap notice — field config saves + round-trips, but the
                output grid only re-renders when the app pivot engine re-applies. */}
            <div
              data-testid="cs-pivot-fields-apply-notice"
              style={{
                display: 'flex',
                gap: 6,
                fontSize: 11,
                lineHeight: 1.4,
                opacity: 0.75,
                background: 'var(--cs-chrome-surface, #f4f6fb)',
                border,
                borderRadius: 6,
                padding: '6px 8px',
                marginBottom: 12,
              }}
            >
              <Icon name="info" size={14} style={{ opacity: 0.7, flexShrink: 0, marginTop: 1 }} />
              <span>
                Field changes are saved to the PivotTable. The output cells refresh once the pivot
                engine re-applies the layout.
              </span>
            </div>

            {/* Source field list */}
            <section style={{ marginBottom: 14 }}>
              <h3 style={sectionTitle}>Choose fields</h3>
              <ul
                data-testid="cs-pivot-fields-list"
                style={{ listStyle: 'none', margin: 0, padding: 0 }}
              >
                {source.headers.map((h, col) => {
                  const ax = axisOf(model, col);
                  const badge =
                    ax === 'rows' ? 'R' : ax === 'cols' ? 'C' : ax === 'filters' ? '▽' : '';
                  return (
                    <li
                      key={col}
                      style={fieldRow}
                      draggable
                      data-testid={`cs-pivot-fields-field-${col}`}
                      onDragStart={(e) => {
                        e.dataTransfer.effectAllowed = 'move';
                        e.dataTransfer.setData(
                          'text/plain',
                          JSON.stringify({ from: 'list', column: col } satisfies DragPayload),
                        );
                      }}
                    >
                      <Icon name="drag_indicator" size={16} style={{ opacity: 0.4 }} />
                      <span
                        aria-hidden
                        style={{
                          width: 8,
                          height: 8,
                          borderRadius: '50%',
                          flexShrink: 0,
                          background: placed.has(col)
                            ? 'var(--cs-chrome-active-fg, #0e7490)'
                            : 'transparent',
                          border: placed.has(col)
                            ? 'none'
                            : '1px solid var(--cs-chrome-border, #c8ccd4)',
                        }}
                      />
                      <span
                        title={h}
                        style={{
                          flex: 1,
                          overflow: 'hidden',
                          textOverflow: 'ellipsis',
                          whiteSpace: 'nowrap',
                        }}
                      >
                        {h}
                      </span>
                      {badge && (
                        <span
                          style={{
                            fontSize: 11,
                            fontWeight: 700,
                            opacity: 0.6,
                            minWidth: 14,
                            textAlign: 'center',
                          }}
                        >
                          {badge}
                        </span>
                      )}
                      <div style={{ position: 'relative' }}>
                        <button
                          type="button"
                          aria-label={`Add ${h} to a zone`}
                          data-testid={`cs-pivot-fields-add-${col}`}
                          onClick={() => setMenuFor((cur) => (cur === col ? null : col))}
                          style={iconBtn}
                        >
                          <Icon name="add" size={16} />
                        </button>
                        {menuFor === col && (
                          <div
                            role="menu"
                            data-testid={`cs-pivot-fields-add-menu-${col}`}
                            style={{
                              position: 'absolute',
                              right: 0,
                              top: '100%',
                              zIndex: 10,
                              minWidth: 150,
                              background: 'var(--cs-chrome-bg, #fff)',
                              border,
                              borderRadius: 6,
                              boxShadow: '0 4px 16px rgba(0,0,0,0.14)',
                              padding: 4,
                            }}
                          >
                            {(['filters', 'rows', 'cols', 'values'] as ZoneId[]).map((zone) => (
                              <button
                                key={zone}
                                type="button"
                                role="menuitem"
                                data-testid={`cs-pivot-fields-add-${col}-${zone}`}
                                onClick={() => {
                                  setMenuFor(null);
                                  commit(addFieldToZone(model, col, zone, optsFor(col)));
                                }}
                                style={{
                                  display: 'block',
                                  width: '100%',
                                  textAlign: 'left',
                                  border: 'none',
                                  background: 'transparent',
                                  cursor: 'pointer',
                                  font: 'inherit',
                                  color: 'inherit',
                                  padding: '6px 8px',
                                  borderRadius: 4,
                                }}
                              >
                                Add to {ZONE_LABELS[zone]}
                              </button>
                            ))}
                          </div>
                        )}
                      </div>
                    </li>
                  );
                })}
                {source.headers.length === 0 && (
                  <li style={{ fontSize: 12, opacity: 0.6, padding: '6px 2px' }}>
                    No source fields found.
                  </li>
                )}
              </ul>
            </section>

            {/* Drop zones */}
            <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
              <Zone zone="filters" model={model} onDrop={onZoneDrop}>
                {(col, i) => {
                  const all = source.distinct(col);
                  const allowed = filterAllowedCount(model, i, all.length);
                  const stored = model.filters?.[i]?.allowedValues ?? [];
                  const isAllowed = (v: string) =>
                    stored.length === 0 ? true : stored.includes(v);
                  const expanded = expandedFilter === i;
                  return (
                    <Chip
                      key={`f-${i}`}
                      label={labelFor(col)}
                      zone="filters"
                      index={i}
                      count={(model.filters ?? []).length}
                      onRemove={() => commit(removeFieldFromZone(model, 'filters', i))}
                      onMove={(dir) => commit(moveWithinZone(model, 'filters', i, i + dir))}
                    >
                      <button
                        type="button"
                        data-testid={`cs-pivot-fields-filter-toggle-${i}`}
                        aria-expanded={expanded}
                        onClick={() => setExpandedFilter(expanded ? null : i)}
                        style={{
                          display: 'flex',
                          alignItems: 'center',
                          gap: 4,
                          border: 'none',
                          background: 'transparent',
                          cursor: 'pointer',
                          color: 'inherit',
                          font: 'inherit',
                          fontSize: 12,
                          padding: 0,
                        }}
                      >
                        <Icon name={expanded ? 'expand_less' : 'expand_more'} size={16} />
                        <span>
                          {allowed} of {all.length} selected
                        </span>
                      </button>
                      {expanded && (
                        <div
                          data-testid={`cs-pivot-fields-filter-values-${i}`}
                          style={{
                            maxHeight: 160,
                            overflow: 'auto',
                            display: 'flex',
                            flexDirection: 'column',
                            gap: 2,
                          }}
                        >
                          <div style={{ display: 'flex', gap: 8, marginBottom: 4 }}>
                            <button
                              type="button"
                              data-testid={`cs-pivot-fields-filter-all-${i}`}
                              onClick={() => commit(setFilterValues(model, i, all))}
                              style={linkBtn}
                            >
                              Select all
                            </button>
                            <button
                              type="button"
                              data-testid={`cs-pivot-fields-filter-clear-${i}`}
                              onClick={() => commit(setFilterValues(model, i, []))}
                              style={linkBtn}
                            >
                              Clear
                            </button>
                          </div>
                          {all.map((v) => (
                            <label
                              key={v}
                              style={{
                                display: 'flex',
                                alignItems: 'center',
                                gap: 6,
                                fontSize: 12,
                              }}
                            >
                              <input
                                type="checkbox"
                                data-testid={`cs-pivot-fields-filter-${i}-${v || 'blank'}`}
                                checked={isAllowed(v)}
                                onChange={(e) =>
                                  commit(toggleFilterValue(model, i, v, e.target.checked, all))
                                }
                              />
                              <span>{v === '' ? '(blank)' : v}</span>
                            </label>
                          ))}
                        </div>
                      )}
                    </Chip>
                  );
                }}
              </Zone>

              <Zone zone="cols" model={model} onDrop={onZoneDrop}>
                {(col, i) => (
                  <Chip
                    key={`c-${i}`}
                    label={labelFor(col)}
                    zone="cols"
                    index={i}
                    count={model.cols.length}
                    onRemove={() => commit(removeFieldFromZone(model, 'cols', i))}
                    onMove={(dir) => commit(moveWithinZone(model, 'cols', i, i + dir))}
                  />
                )}
              </Zone>

              <Zone zone="rows" model={model} onDrop={onZoneDrop}>
                {(col, i) => (
                  <Chip
                    key={`r-${i}`}
                    label={labelFor(col)}
                    zone="rows"
                    index={i}
                    count={model.rows.length}
                    onRemove={() => commit(removeFieldFromZone(model, 'rows', i))}
                    onMove={(dir) => commit(moveWithinZone(model, 'rows', i, i + dir))}
                  >
                    <select
                      aria-label="Group dates by"
                      data-testid={`cs-pivot-fields-rows-grouping-${i}`}
                      value={model.rows[i]?.grouping ?? 'none'}
                      onChange={(e) =>
                        commit(updateRowGrouping(model, i, e.target.value as DateGrouping))
                      }
                      style={selectStyle}
                    >
                      {(Object.keys(PIVOT_DATE_GROUP_LABELS) as DateGrouping[]).map((g) => (
                        <option key={g} value={g}>
                          {PIVOT_DATE_GROUP_LABELS[g]}
                        </option>
                      ))}
                    </select>
                  </Chip>
                )}
              </Zone>

              <Zone zone="values" model={model} onDrop={onZoneDrop}>
                {(col, i) => (
                  <Chip
                    key={`v-${i}`}
                    label={`${PIVOT_AGG_LABELS[model.values[i].agg]} of ${labelFor(col)}`}
                    zone="values"
                    index={i}
                    count={model.values.length}
                    onRemove={
                      hasValues(model) && model.values.length > 1
                        ? () => commit(removeFieldFromZone(model, 'values', i))
                        : undefined
                    }
                    onMove={(dir) => commit(moveWithinZone(model, 'values', i, i + dir))}
                  >
                    <div style={{ display: 'flex', gap: 4 }}>
                      <select
                        aria-label="Summarize values by"
                        data-testid={`cs-pivot-fields-values-agg-${i}`}
                        value={model.values[i].agg}
                        onChange={(e) =>
                          commit(
                            updateValueField(model, i, { agg: e.target.value as PivotAggregation }),
                          )
                        }
                        style={selectStyle}
                      >
                        {(Object.keys(PIVOT_AGG_LABELS) as PivotAggregation[]).map((a) => (
                          <option key={a} value={a}>
                            {PIVOT_AGG_LABELS[a]}
                          </option>
                        ))}
                      </select>
                      <select
                        aria-label="Show values as"
                        data-testid={`cs-pivot-fields-values-showas-${i}`}
                        value={model.values[i].showAs ?? 'normal'}
                        onChange={(e) =>
                          commit(
                            updateValueField(model, i, { showAs: e.target.value as PivotShowAs }),
                          )
                        }
                        style={selectStyle}
                      >
                        {(Object.keys(PIVOT_SHOW_AS_LABELS) as PivotShowAs[]).map((s) => (
                          <option key={s} value={s}>
                            {PIVOT_SHOW_AS_LABELS[s]}
                          </option>
                        ))}
                      </select>
                    </div>
                  </Chip>
                )}
              </Zone>
            </div>
          </>
        )}
      </div>
    </div>
  );
}

const linkBtn: CSSProperties = {
  border: 'none',
  background: 'transparent',
  cursor: 'pointer',
  color: 'var(--cs-chrome-active-fg, #0e7490)',
  font: 'inherit',
  fontSize: 12,
  padding: 0,
};

/* ------------------------------------------------------------------ *
 * Zone — one drop target (Filters / Columns / Rows / Values).
 * ------------------------------------------------------------------ */

function Zone({
  zone,
  model,
  children,
  onDrop,
}: {
  zone: ZoneId;
  model: PivotModel;
  children: (column: number, index: number) => ReactNode;
  onDrop: (zone: ZoneId, raw: string) => void;
}) {
  const [over, setOver] = useState(false);
  const entries: number[] =
    zone === 'values'
      ? model.values.map((v) => v.column)
      : zone === 'filters'
        ? (model.filters ?? []).map((f) => f.column)
        : model[zone].map((e) => e.column);

  return (
    <section
      data-testid={`cs-pivot-fields-zone-${zone}`}
      onDragOver={(e) => {
        e.preventDefault();
        e.dataTransfer.dropEffect = 'move';
        if (!over) setOver(true);
      }}
      onDragLeave={(e) => {
        if (!e.currentTarget.contains(e.relatedTarget as Node)) setOver(false);
      }}
      onDrop={(e) => {
        e.preventDefault();
        setOver(false);
        onDrop(zone, e.dataTransfer.getData('text/plain'));
      }}
      style={{
        border: over
          ? '1px dashed var(--cs-chrome-active-fg, #0e7490)'
          : '1px dashed var(--cs-chrome-border, #d5d9e0)',
        borderRadius: 8,
        padding: 8,
        background: over ? 'var(--cs-chrome-surface, #eef6f8)' : 'transparent',
        minHeight: 56,
      }}
    >
      <h4 style={{ ...sectionTitle, margin: '0 0 6px' }}>{ZONE_LABELS[zone]}</h4>
      <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
        {entries.length === 0 ? (
          <div style={{ fontSize: 12, opacity: 0.5, padding: '4px 2px' }}>Drop or add fields</div>
        ) : (
          entries.map((col, i) => children(col, i))
        )}
      </div>
    </section>
  );
}

/* ------------------------------------------------------------------ *
 * Chip — a placed field: label, move up/down, remove, extra controls.
 * ------------------------------------------------------------------ */

function Chip({
  label,
  zone,
  index,
  count,
  onRemove,
  onMove,
  children,
}: {
  label: string;
  zone: ZoneId;
  index: number;
  count: number;
  onRemove?: () => void;
  onMove: (dir: -1 | 1) => void;
  children?: ReactNode;
}) {
  return (
    <div
      data-testid={`cs-pivot-fields-chip-${zone}-${index}`}
      draggable
      onDragStart={(e) => {
        e.dataTransfer.effectAllowed = 'move';
        e.dataTransfer.setData(
          'text/plain',
          JSON.stringify({ from: 'zone', zone, index } satisfies DragPayload),
        );
      }}
      style={{ ...chipStyle, cursor: 'grab' }}
    >
      <div style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
        <Icon name="drag_indicator" size={16} style={{ opacity: 0.4 }} />
        <span
          title={label}
          style={{
            flex: 1,
            fontSize: 13,
            overflow: 'hidden',
            textOverflow: 'ellipsis',
            whiteSpace: 'nowrap',
          }}
        >
          {label}
        </span>
        {index > 0 && (
          <button
            type="button"
            aria-label="Move up"
            data-testid={`cs-pivot-fields-chip-${zone}-${index}-up`}
            onClick={() => onMove(-1)}
            style={iconBtn}
          >
            <Icon name="keyboard_arrow_up" size={16} />
          </button>
        )}
        {index < count - 1 && (
          <button
            type="button"
            aria-label="Move down"
            data-testid={`cs-pivot-fields-chip-${zone}-${index}-down`}
            onClick={() => onMove(1)}
            style={iconBtn}
          >
            <Icon name="keyboard_arrow_down" size={16} />
          </button>
        )}
        {onRemove && (
          <button
            type="button"
            aria-label={`Remove ${label}`}
            data-testid={`cs-pivot-fields-chip-${zone}-${index}-remove`}
            onClick={onRemove}
            style={iconBtn}
          >
            <Icon name="close" size={16} />
          </button>
        )}
      </div>
      {children}
    </div>
  );
}
