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

import { useEffect, useMemo, useState } from 'react';
import type { FUniver } from '@univerjs/core/facade';
import { useUniverAPI } from '../use-univer';
import { useUI } from '../use-ui';
import { Icon } from '../shell/Icon';
import { usePivots } from './pivots-context';
import { applyPivot } from './apply';
import {
  PIVOT_AGG_LABELS,
  PIVOT_DATE_GROUP_LABELS,
  PIVOT_SHOW_AS_LABELS,
  type DateGrouping,
  type PivotAggregation,
  type PivotModel,
  type PivotShowAs,
} from './types';
import {
  ZONE_LABELS,
  addFieldToZone,
  axisOf,
  filterAllowedCount,
  hasValues,
  moveWithinZone,
  placedColumns,
  removeFieldFromZone,
  setFilterValues,
  toggleFilterValue,
  updateRowGrouping,
  updateValueField,
  type ZoneId,
} from './fields-model';

/**
 * Excel's "PivotTable Fields" task pane. Lists the source fields and the
 * four drop zones — Filters / Columns / Rows / Values — and re-applies
 * the pivot live as the user reconfigures it, so you don't have to delete
 * and re-insert to change the layout.
 *
 * Field assignment uses click-to-assign (each field's "+" opens the
 * four-zone menu, mirroring Excel's right-click "Add to Row Labels / …"
 * affordance) plus per-chip remove / reorder, Values agg + Show-Values-As
 * editing, Rows date-grouping, and a per-value checklist on report filters
 * (Filters zone) that actually narrows the source records. Drag-and-drop
 * between zones is the remaining slice.
 */

type SourceView = {
  headers: string[];
  /** Distinct string values in a source column (for report filters). */
  distinct: (col: number) => string[];
  /** Best-guess numeric column → default a new value field to Sum vs Count. */
  isNumeric: (col: number) => boolean;
};

function readSource(api: FUniver | null, model: PivotModel | null): SourceView {
  const empty: SourceView = { headers: [], distinct: () => [], isNumeric: () => false };
  if (!api || !model) return empty;
  const wb = api.getActiveWorkbook();
  if (!wb) return empty;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
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

export function PivotFieldsPanel() {
  const api = useUniverAPI();
  const ui = useUI();
  const pivots = usePivots();

  const [selectedId, setSelectedId] = useState<string | null>(null);
  // Field-list "+" menu — which source column's zone picker is open.
  const [menuFor, setMenuFor] = useState<number | null>(null);
  // Filters zone — which report filter's value checklist is expanded.
  const [expandedFilter, setExpandedFilter] = useState<number | null>(null);

  // Keep a valid selection: default to the most-recently-inserted pivot,
  // and recover if the selected pivot was deleted.
  useEffect(() => {
    if (pivots.pivots.length === 0) {
      if (selectedId !== null) setSelectedId(null);
      return;
    }
    if (!selectedId || !pivots.pivots.some((p) => p.id === selectedId)) {
      setSelectedId(pivots.pivots[pivots.pivots.length - 1].id);
    }
  }, [pivots.pivots, selectedId]);

  const model = useMemo(
    () => pivots.pivots.find((p) => p.id === selectedId) ?? null,
    [pivots.pivots, selectedId],
  );
  const source = useMemo(() => readSource(api, model), [api, model]);

  // Re-apply the edited model to the sheet and persist it on the store.
  const commit = (next: PivotModel) => {
    if (!api) return;
    const extent = applyPivot(api, next, next.lastOutputExtent ?? null);
    pivots.update(next.id, {
      rows: next.rows,
      cols: next.cols,
      values: next.values,
      filters: next.filters,
      lastOutputExtent: extent ?? next.lastOutputExtent,
    });
  };

  const placed = model ? placedColumns(model) : new Set<number>();

  return (
    <aside className="side-panel pivot-fields-panel" data-testid="pivot-fields-panel">
      <header className="side-panel__header">
        <Icon name="pivot_table_chart" size="sm" />
        <h2 className="side-panel__title">PivotTable Fields</h2>
        <button
          type="button"
          className="side-panel__close"
          aria-label="Close PivotTable Fields panel"
          onClick={ui.togglePivotPanel}
        >
          <Icon name="close" size="sm" />
        </button>
      </header>

      <div className="side-panel__body pivot-fields-panel__body">
        {!model ? (
          <div className="side-panel__empty" data-testid="pivot-fields-empty">
            <Icon name="pivot_table_chart" size="lg" className="side-panel__empty-icon" />
            <div className="side-panel__empty-title">No PivotTable selected</div>
            <div className="side-panel__empty-body">
              Insert a PivotTable, then configure its fields here.
            </div>
            <button
              type="button"
              className="btn-primary side-panel__empty-cta"
              data-testid="pivot-fields-insert-cta"
              onClick={() => document.dispatchEvent(new CustomEvent('casual-open-insert-pivot'))}
            >
              Insert PivotTable
            </button>
          </div>
        ) : (
          <>
            {pivots.pivots.length > 1 && (
              <label className="pivot-fields-panel__picker">
                <span>PivotTable</span>
                <select
                  data-testid="pivot-fields-picker"
                  value={selectedId ?? ''}
                  onChange={(e) => setSelectedId(e.target.value)}
                >
                  {pivots.pivots.map((p, i) => (
                    <option key={p.id} value={p.id}>
                      {p.title ?? `PivotTable ${i + 1}`}
                    </option>
                  ))}
                </select>
              </label>
            )}

            {/* Source field list */}
            <section className="pivot-fields-panel__section">
              <h3 className="pivot-fields-panel__section-title">Choose fields</h3>
              <ul className="pivot-fields-panel__field-list" data-testid="pivot-fields-list">
                {source.headers.map((h, col) => {
                  const ax = model ? axisOf(model, col) : null;
                  const badge =
                    ax === 'rows' ? 'R' : ax === 'cols' ? 'C' : ax === 'filters' ? '▽' : '';
                  return (
                    <li key={col} className="pivot-fields-panel__field">
                      <span
                        className={`pivot-fields-panel__field-dot${
                          placed.has(col) ? ' pivot-fields-panel__field-dot--on' : ''
                        }`}
                        aria-hidden
                      />
                      <span className="pivot-fields-panel__field-name" title={h}>
                        {h}
                      </span>
                      {badge && <span className="pivot-fields-panel__field-badge">{badge}</span>}
                      <div className="pivot-fields-panel__field-add">
                        <button
                          type="button"
                          className="pivot-fields-panel__add-btn"
                          aria-label={`Add ${h} to a zone`}
                          data-testid={`pivot-fields-add-${col}`}
                          onClick={() => setMenuFor((cur) => (cur === col ? null : col))}
                        >
                          <Icon name="add" size="sm" />
                        </button>
                        {menuFor === col && (
                          <div
                            className="pivot-fields-panel__add-menu"
                            data-testid={`pivot-fields-add-menu-${col}`}
                            role="menu"
                          >
                            {(['filters', 'rows', 'cols', 'values'] as ZoneId[]).map((zone) => (
                              <button
                                key={zone}
                                type="button"
                                role="menuitem"
                                data-testid={`pivot-fields-add-${col}-${zone}`}
                                onClick={() => {
                                  if (!model) return;
                                  const next =
                                    zone === 'values'
                                      ? addFieldToZone(model, col, 'values', {
                                          defaultAgg: source.isNumeric(col) ? 'sum' : 'count',
                                        })
                                      : zone === 'filters'
                                        ? addFieldToZone(model, col, 'filters', {
                                            allowedValues: source.distinct(col),
                                          })
                                        : addFieldToZone(model, col, zone);
                                  setMenuFor(null);
                                  commit(next);
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
              </ul>
            </section>

            {/* Drop zones */}
            <div className="pivot-fields-panel__zones">
              <Zone
                zone="filters"
                model={model}
                renderChip={(col, i) => {
                  const all = source.distinct(col);
                  const allowed = filterAllowedCount(model, i, all.length);
                  const stored = model.filters?.[i]?.allowedValues ?? [];
                  // Empty stored list = "all pass" (slice-1 convention).
                  const isAllowed = (v: string) =>
                    stored.length === 0 ? true : stored.includes(v);
                  const expanded = expandedFilter === i;
                  return (
                    <Chip
                      key={`f-${i}`}
                      label={source.headers[col] ?? `Column ${col + 1}`}
                      zone="filters"
                      index={i}
                      count={(model.filters ?? []).length}
                      onRemove={() => commit(removeFieldFromZone(model, 'filters', i))}
                      onMove={(dir) => commit(moveWithinZone(model, 'filters', i, i + dir))}
                    >
                      <button
                        type="button"
                        className="pivot-fields-panel__filter-toggle"
                        data-testid={`pivot-fields-filter-toggle-${i}`}
                        aria-expanded={expanded}
                        onClick={() => setExpandedFilter(expanded ? null : i)}
                      >
                        <Icon name={expanded ? 'expand_less' : 'expand_more'} size="sm" />
                        <span>
                          {allowed} of {all.length} selected
                        </span>
                      </button>
                      {expanded && (
                        <div
                          className="pivot-fields-panel__filter-values"
                          data-testid={`pivot-fields-filter-values-${i}`}
                        >
                          <div className="pivot-fields-panel__filter-bulk">
                            <button
                              type="button"
                              data-testid={`pivot-fields-filter-all-${i}`}
                              onClick={() => commit(setFilterValues(model, i, all))}
                            >
                              Select all
                            </button>
                            <button
                              type="button"
                              data-testid={`pivot-fields-filter-clear-${i}`}
                              onClick={() => commit(setFilterValues(model, i, []))}
                            >
                              Clear
                            </button>
                          </div>
                          {all.map((v) => (
                            <label key={v} className="pivot-fields-panel__filter-value">
                              <input
                                type="checkbox"
                                data-testid={`pivot-fields-filter-${i}-${v || 'blank'}`}
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
              />
              <Zone
                zone="cols"
                model={model}
                renderChip={(col, i) => (
                  <Chip
                    key={`c-${i}`}
                    label={source.headers[col] ?? `Column ${col + 1}`}
                    zone="cols"
                    index={i}
                    count={model.cols.length}
                    onRemove={() => commit(removeFieldFromZone(model, 'cols', i))}
                    onMove={(dir) => commit(moveWithinZone(model, 'cols', i, i + dir))}
                  />
                )}
              />
              <Zone
                zone="rows"
                model={model}
                renderChip={(col, i) => (
                  <Chip
                    key={`r-${i}`}
                    label={source.headers[col] ?? `Column ${col + 1}`}
                    zone="rows"
                    index={i}
                    count={model.rows.length}
                    onRemove={() => commit(removeFieldFromZone(model, 'rows', i))}
                    onMove={(dir) => commit(moveWithinZone(model, 'rows', i, i + dir))}
                  >
                    <select
                      className="pivot-fields-panel__chip-select"
                      aria-label="Group dates by"
                      data-testid={`pivot-fields-rows-grouping-${i}`}
                      value={model.rows[i]?.grouping ?? 'none'}
                      onChange={(e) =>
                        commit(updateRowGrouping(model, i, e.target.value as DateGrouping))
                      }
                    >
                      {(Object.keys(PIVOT_DATE_GROUP_LABELS) as DateGrouping[]).map((g) => (
                        <option key={g} value={g}>
                          {PIVOT_DATE_GROUP_LABELS[g]}
                        </option>
                      ))}
                    </select>
                  </Chip>
                )}
              />
              <Zone
                zone="values"
                model={model}
                renderChip={(col, i) => (
                  <Chip
                    key={`v-${i}`}
                    label={`${PIVOT_AGG_LABELS[model.values[i].agg]} of ${
                      source.headers[col] ?? `Column ${col + 1}`
                    }`}
                    zone="values"
                    index={i}
                    count={model.values.length}
                    // Block removing the last value field — a value-less pivot
                    // renders nothing.
                    onRemove={
                      hasValues(model) && model.values.length > 1
                        ? () => commit(removeFieldFromZone(model, 'values', i))
                        : undefined
                    }
                    onMove={(dir) => commit(moveWithinZone(model, 'values', i, i + dir))}
                  >
                    <div className="pivot-fields-panel__chip-controls">
                      <select
                        className="pivot-fields-panel__chip-select"
                        aria-label="Summarize values by"
                        data-testid={`pivot-fields-values-agg-${i}`}
                        value={model.values[i].agg}
                        onChange={(e) =>
                          commit(
                            updateValueField(model, i, {
                              agg: e.target.value as PivotAggregation,
                            }),
                          )
                        }
                      >
                        {(Object.keys(PIVOT_AGG_LABELS) as PivotAggregation[]).map((a) => (
                          <option key={a} value={a}>
                            {PIVOT_AGG_LABELS[a]}
                          </option>
                        ))}
                      </select>
                      <select
                        className="pivot-fields-panel__chip-select"
                        aria-label="Show values as"
                        data-testid={`pivot-fields-values-showas-${i}`}
                        value={model.values[i].showAs ?? 'normal'}
                        onChange={(e) =>
                          commit(
                            updateValueField(model, i, {
                              showAs: e.target.value as PivotShowAs,
                            }),
                          )
                        }
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
              />
            </div>
          </>
        )}
      </div>
    </aside>
  );
}

/** One drop zone (Filters / Columns / Rows / Values) with its chips. */
function Zone({
  zone,
  model,
  renderChip,
}: {
  zone: ZoneId;
  model: PivotModel;
  renderChip: (column: number, index: number) => React.ReactNode;
}) {
  const entries: number[] =
    zone === 'values'
      ? model.values.map((v) => v.column)
      : zone === 'filters'
        ? (model.filters ?? []).map((f) => f.column)
        : model[zone].map((e) => e.column);

  return (
    <section className="pivot-fields-panel__zone" data-testid={`pivot-fields-zone-${zone}`}>
      <h4 className="pivot-fields-panel__zone-title">{ZONE_LABELS[zone]}</h4>
      <div className="pivot-fields-panel__zone-body">
        {entries.length === 0 ? (
          <div className="pivot-fields-panel__zone-empty">Drop or add fields</div>
        ) : (
          entries.map((col, i) => renderChip(col, i))
        )}
      </div>
    </section>
  );
}

/** A placed-field chip: label, optional extra controls, move + remove. */
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
  children?: React.ReactNode;
}) {
  return (
    <div className="pivot-fields-panel__chip" data-testid={`pivot-fields-chip-${zone}-${index}`}>
      <div className="pivot-fields-panel__chip-head">
        <span className="pivot-fields-panel__chip-label" title={label}>
          {label}
        </span>
        <div className="pivot-fields-panel__chip-actions">
          {index > 0 && (
            <button
              type="button"
              className="pivot-fields-panel__chip-btn"
              aria-label="Move up"
              data-testid={`pivot-fields-chip-${zone}-${index}-up`}
              onClick={() => onMove(-1)}
            >
              <Icon name="keyboard_arrow_up" size="sm" />
            </button>
          )}
          {index < count - 1 && (
            <button
              type="button"
              className="pivot-fields-panel__chip-btn"
              aria-label="Move down"
              data-testid={`pivot-fields-chip-${zone}-${index}-down`}
              onClick={() => onMove(1)}
            >
              <Icon name="keyboard_arrow_down" size="sm" />
            </button>
          )}
          {onRemove && (
            <button
              type="button"
              className="pivot-fields-panel__chip-btn pivot-fields-panel__chip-btn--remove"
              aria-label={`Remove ${label}`}
              data-testid={`pivot-fields-chip-${zone}-${index}-remove`}
              onClick={onRemove}
            >
              <Icon name="close" size="sm" />
            </button>
          )}
        </div>
      </div>
      {children}
    </div>
  );
}
