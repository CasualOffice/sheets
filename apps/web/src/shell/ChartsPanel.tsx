import { useMemo, useState } from 'react';
import { Icon } from './Icon';
import { useUI } from '../use-ui';
import { useUniverAPI } from '../use-univer';
import { useCharts } from '../charts/charts-context';
import {
  getActiveSelectionRange,
  rangeToA1,
  buildChartModelForRange,
} from '../charts/insert-chart';
import { nextChartName } from '../charts/naming';
import { InsertChartDialog } from '../charts/InsertChartDialog';
import {
  CHART_FAMILY_OF,
  CHART_TYPE_LABEL,
  type ChartFamily,
  type ChartModel,
} from '../charts/types';

/**
 * Right-side Charts panel. Equivalent of Excel's Selection Pane scoped
 * to charts on the active sheet: list every chart, click the name to
 * rename, click the source-range badge to flash that range in the grid,
 * delete from the row, and "Insert chart" from the empty-state CTA.
 *
 * Only charts on the active sheet are shown — same scoping Excel uses
 * for its Selection Pane (it switches with the active sheet tab).
 */
const FAMILY_ICONS: Record<ChartFamily, string> = {
  column: 'bar_chart',
  bar: 'align_horizontal_left',
  line: 'show_chart',
  area: 'area_chart',
  pie: 'pie_chart',
  scatter: 'scatter_plot',
};

export function ChartsPanel() {
  const ui = useUI();
  const api = useUniverAPI();
  const { charts, insert, remove, update } = useCharts();
  const [renaming, setRenaming] = useState<{ id: string; draft: string } | null>(null);
  const [showInsert, setShowInsert] = useState(false);
  const [insertDefault, setInsertDefault] = useState('A1');

  const activeSheetId = useMemo(() => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api?.getActiveWorkbook?.()?.getActiveSheet();
    return ws?.getSheetId?.() ?? null;
  }, [api, charts]);

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
    if (!api) return;
    const sel = getActiveSelectionRange(api);
    setInsertDefault(sel ? rangeToA1(sel) : 'A1');
    setShowInsert(true);
  };

  return (
    <aside className="side-panel charts-panel" data-testid="charts-panel">
      <header className="side-panel__header">
        <Icon name="bar_chart" size="sm" />
        <h2 className="side-panel__title">Charts</h2>
        {!empty && <span className="side-panel__count">{visible.length}</span>}
        <button
          type="button"
          className="side-panel__close"
          aria-label="Close charts panel"
          onClick={ui.toggleChartsPanel}
        >
          <Icon name="close" size="sm" />
        </button>
      </header>
      <div className="charts-panel__body">
        {empty ? (
          <div className="charts-panel__empty" data-testid="charts-panel-empty">
            <Icon name="bar_chart" size="lg" className="charts-panel__empty-icon" />
            <div className="charts-panel__empty-title">No charts on this sheet</div>
            <div className="charts-panel__empty-body">
              Select the data range you want to plot, then click below — or use{' '}
              <strong>Insert → Chart</strong> from the menu.
            </div>
            <button
              type="button"
              className="btn-primary charts-panel__empty-cta"
              data-testid="charts-panel-empty-cta"
              disabled={!api}
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
                disabled={!api}
                onClick={openInsert}
              >
                <Icon name="add" /> Insert chart
              </button>
            </li>
          </ul>
        )}
      </div>

      {showInsert && api && (
        <InsertChartDialog
          api={api}
          defaultSourceA1={insertDefault}
          onCancel={() => setShowInsert(false)}
          onConfirm={({ source, type }) => {
            const model = buildChartModelForRange(api, source, type);
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
