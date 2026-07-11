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

import { useMemo, useState } from 'react';
import type { PanelComponentProps } from '../chrome/extensions';
import { Icon } from '../chrome/Icon';
import { useCharts } from './charts-context';
import { getActiveSelectionRange, rangeToA1, buildChartModelForRange } from './insert-chart';
import { nextChartName } from './naming';
import { InsertChartDialog } from './InsertChartDialog';
import { CHART_FAMILY_OF, CHART_TYPE_LABEL, type ChartFamily, type ChartModel } from './types';

/**
 * Right-side Charts panel. Equivalent of Excel's Selection Pane scoped
 * to charts on the active sheet: list every chart, click the name to
 * rename, click the source-range badge to flash that range in the grid,
 * delete from the row, and "Insert chart" from the empty-state CTA.
 *
 * Only charts on the active sheet are shown — same scoping Excel uses
 * for its Selection Pane (it switches with the active sheet tab).
 *
 * SDK port: this is a chrome side-panel (`{ api, onClose }`). The app's
 * `useUI().toggleChartsPanel` becomes the panel host's `onClose`; the
 * FUniver facade is reached through `api.univer`. Chart data still comes
 * from `useCharts()` (mount `<ChartsProvider api={...}>` above the chrome).
 */
const FAMILY_ICONS: Record<ChartFamily, string> = {
  column: 'bar_chart',
  bar: 'align_horizontal_left',
  line: 'show_chart',
  area: 'area_chart',
  pie: 'pie_chart',
  scatter: 'scatter_plot',
};

export function ChartsPanel({ api, onClose }: PanelComponentProps) {
  // FUniver facade — reached through the SDK handle.
  const univer = api.univer;
  const { charts, insert, remove, update } = useCharts();
  const [renaming, setRenaming] = useState<{ id: string; draft: string } | null>(null);
  const [showInsert, setShowInsert] = useState(false);
  const [insertDefault, setInsertDefault] = useState('A1');

  const activeSheetId = useMemo(() => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = univer?.getActiveWorkbook?.()?.getActiveSheet();
    return ws?.getSheetId?.() ?? null;
  }, [univer, charts]);

  const visible = useMemo<ChartModel[]>(
    () => (activeSheetId ? charts.filter((c) => c.sheetId === activeSheetId) : charts),
    [charts, activeSheetId],
  );
  const empty = visible.length === 0;

  const onRenameCommit = (id: string, prev: string) => {
    if (!renaming || renaming.id !== id) return;
    const draft = renaming.draft.trim();
    if (!draft || draft === prev) {
      setRenaming(null);
      return;
    }
    update(id, { title: draft });
    setRenaming(null);
  };

  const openInsert = () => {
    if (!univer) return;
    const sel = getActiveSelectionRange(univer);
    setInsertDefault(sel ? rangeToA1(sel) : 'A1');
    setShowInsert(true);
  };

  return (
    <aside className="side-panel charts-panel" data-testid="charts-panel">
      <header className="side-panel__header">
        <Icon name="bar_chart" size={16} />
        <h2 className="side-panel__title">Charts</h2>
        {!empty && <span className="side-panel__count">{visible.length}</span>}
        <button
          type="button"
          className="side-panel__close"
          aria-label="Close charts panel"
          onClick={onClose}
        >
          <Icon name="close" size={16} />
        </button>
      </header>
      <div className="charts-panel__body">
        {empty ? (
          <div className="charts-panel__empty" data-testid="charts-panel-empty">
            <Icon name="bar_chart" size={32} style={{ opacity: 0.4 }} />
            <div className="charts-panel__empty-title">No charts on this sheet</div>
            <div className="charts-panel__empty-body">
              Select the data range you want to plot, then click below — or use{' '}
              <strong>Insert → Chart</strong> from the menu.
            </div>
            <button
              type="button"
              className="btn-primary charts-panel__empty-cta"
              data-testid="charts-panel-empty-cta"
              disabled={!univer}
              onClick={openInsert}
            >
              Insert chart
            </button>
          </div>
        ) : (
          <ul className="charts-panel__list">
            {visible.map((c) => {
              const isRenaming = renaming?.id === c.id;
              const displayName = c.title ?? 'Chart';
              return (
                <li
                  className="charts-panel__row"
                  key={c.id}
                  data-testid={`charts-panel-row-${c.id}`}
                >
                  <div className="charts-panel__name">
                    <span
                      className="material-symbols-outlined charts-panel__type-icon"
                      aria-hidden="true"
                    >
                      {FAMILY_ICONS[CHART_FAMILY_OF[c.type]]}
                    </span>
                    {isRenaming ? (
                      <input
                        autoFocus
                        className="charts-panel__name-input"
                        value={renaming.draft}
                        onChange={(e) => setRenaming({ id: c.id, draft: e.target.value })}
                        onBlur={() => onRenameCommit(c.id, displayName)}
                        onKeyDown={(e) => {
                          if (e.key === 'Enter') onRenameCommit(c.id, displayName);
                          if (e.key === 'Escape') setRenaming(null);
                        }}
                      />
                    ) : (
                      <button
                        type="button"
                        className="charts-panel__name-btn"
                        onClick={() => setRenaming({ id: c.id, draft: displayName })}
                        title="Click to rename"
                      >
                        {displayName}
                      </button>
                    )}
                  </div>
                  <div className="charts-panel__meta">
                    <span className="charts-panel__type-label">{CHART_TYPE_LABEL[c.type]}</span>
                    <span className="charts-panel__range" title="Source range">
                      {rangeToA1(c.source)}
                    </span>
                  </div>
                  <button
                    type="button"
                    className="charts-panel__delete"
                    aria-label={`Delete ${displayName}`}
                    title="Delete chart"
                    onClick={() => remove(c.id)}
                  >
                    <Icon name="delete" />
                  </button>
                </li>
              );
            })}
            <li className="charts-panel__add-row">
              <button
                type="button"
                className="btn-secondary charts-panel__add"
                data-testid="charts-panel-add"
                disabled={!univer}
                onClick={openInsert}
              >
                <Icon name="add" /> Insert chart
              </button>
            </li>
          </ul>
        )}
      </div>

      {showInsert && univer && (
        <InsertChartDialog
          api={univer}
          defaultSourceA1={insertDefault}
          onCancel={() => setShowInsert(false)}
          onConfirm={({ source, type }) => {
            const model = buildChartModelForRange(univer, source, type);
            if (model) {
              insert({ ...model, title: nextChartName(charts) });
            }
            setShowInsert(false);
          }}
        />
      )}
    </aside>
  );
}
