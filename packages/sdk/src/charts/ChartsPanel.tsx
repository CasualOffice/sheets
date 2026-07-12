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
import { PanelHeader, PanelEmpty, IconButton } from '../chrome/panel-shell';

const MUTED = 'var(--color-text-secondary, var(--cs-chrome-muted, #605e5c))';
const DIVIDER = 'var(--color-divider, var(--cs-chrome-border, #edeff3))';
const rowStyle = {
  border: `1px solid ${DIVIDER}`,
  borderRadius: 8,
  padding: 10,
  display: 'flex',
  flexDirection: 'column' as const,
  gap: 6,
};
const nameBtnStyle = {
  flex: 1,
  textAlign: 'left' as const,
  border: 'none',
  background: 'transparent',
  cursor: 'pointer',
  font: 'inherit',
  fontWeight: 600,
  color: 'inherit',
};
const ctaStyle = {
  border: `1px solid ${DIVIDER}`,
  borderRadius: 6,
  padding: '6px 12px',
  cursor: 'pointer',
  font: 'inherit',
  fontWeight: 600,
  display: 'inline-flex',
  alignItems: 'center',
  gap: 6,
  background: 'var(--color-surface, var(--cs-chrome-input-bg, #fff))',
  color: 'var(--color-accent, var(--cs-chrome-active-fg, #0e7490))',
};
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
    <div data-testid="charts-panel" style={{ height: '100%', display: 'flex', flexDirection: 'column' }}>
      <PanelHeader icon="bar_chart" title="Charts" count={visible.length} onClose={onClose} />
      <div style={{ flex: 1, overflow: 'auto', padding: 12 }}>
        {empty ? (
          <PanelEmpty icon="bar_chart" title="No charts on this sheet" testId="charts-panel-empty">
            Select the data range you want to plot, then use <strong>Insert → Chart</strong> — or:
            <div style={{ marginTop: 12 }}>
              <button
                type="button"
                data-testid="charts-panel-empty-cta"
                disabled={!univer}
                onClick={openInsert}
                style={ctaStyle}
              >
                <Icon name="add" size={16} /> Insert chart
              </button>
            </div>
          </PanelEmpty>
        ) : (
          <ul
            style={{
              listStyle: 'none',
              margin: 0,
              padding: 0,
              display: 'flex',
              flexDirection: 'column',
              gap: 8,
            }}
          >
            {visible.map((c) => {
              const isRenaming = renaming?.id === c.id;
              const displayName = c.title ?? 'Chart';
              return (
                <li key={c.id} data-testid={`charts-panel-row-${c.id}`} style={rowStyle}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                    <span
                      className="material-symbols-outlined"
                      aria-hidden="true"
                      style={{ fontSize: 18, color: MUTED, flex: '0 0 auto' }}
                    >
                      {FAMILY_ICONS[CHART_FAMILY_OF[c.type]]}
                    </span>
                    {isRenaming ? (
                      <input
                        autoFocus
                        value={renaming.draft}
                        onChange={(e) => setRenaming({ id: c.id, draft: e.target.value })}
                        onBlur={() => onRenameCommit(c.id, displayName)}
                        onKeyDown={(e) => {
                          if (e.key === 'Enter') onRenameCommit(c.id, displayName);
                          if (e.key === 'Escape') setRenaming(null);
                        }}
                        style={{ flex: 1, font: 'inherit', padding: '2px 4px' }}
                      />
                    ) : (
                      <button
                        type="button"
                        onClick={() => setRenaming({ id: c.id, draft: displayName })}
                        title="Click to rename"
                        style={nameBtnStyle}
                      >
                        {displayName}
                      </button>
                    )}
                    <IconButton
                      name="delete"
                      label={`Delete ${displayName}`}
                      onClick={() => remove(c.id)}
                      size={16}
                    />
                  </div>
                  <div style={{ display: 'flex', gap: 8, fontSize: 12, color: MUTED }}>
                    <span>{CHART_TYPE_LABEL[c.type]}</span>
                    <span title="Source range">{rangeToA1(c.source)}</span>
                  </div>
                </li>
              );
            })}
            <li style={{ listStyle: 'none', marginTop: 4 }}>
              <button
                type="button"
                data-testid="charts-panel-add"
                disabled={!univer}
                onClick={openInsert}
                style={ctaStyle}
              >
                <Icon name="add" size={16} /> Insert chart
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
    </div>
  );
}
