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

import { useEffect, useRef, useState } from 'react';
import type { FUniver } from '@univerjs/core/facade';
import { Icon } from '../chrome/Icon';
import { useCharts } from './charts-context';
import { InsertChartDialog } from './InsertChartDialog';
import { FormatChartDialog } from './FormatChartDialog';

/**
 * Right-click context menu for a selected chart. Mirrors Excel's
 * chart context menu, scoped to the actions we support today:
 *
 *   - Change Chart Type → reopens the Insert dialog seeded with the
 *     chart's current source range and type.
 *   - Rename → in-place edit (focus jumps to the panel rename if it's
 *     open, otherwise inline prompt — for v0.1.1 we just toggle a
 *     small inline input here).
 *   - Delete → drops the chart from the store.
 *
 * Closes on outside click or Escape.
 */
type Props = {
  chartId: string;
  x: number;
  y: number;
  onClose: () => void;
};

export function ChartContextMenu({ chartId, x, y, onClose }: Props) {
  const { api: sdkApi, charts, remove, update } = useCharts();
  // FUniver facade — reached through the SDK handle.
  const api = sdkApi.univer;
  const chart = charts.find((c) => c.id === chartId);
  const menuRef = useRef<HTMLDivElement>(null);
  const [renaming, setRenaming] = useState<string | null>(null);
  const [showChangeType, setShowChangeType] = useState(false);
  const [showFormat, setShowFormat] = useState(false);
  const dialogOpen = showChangeType || showFormat;

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        if (dialogOpen) return; // dialog handles its own Escape
        onClose();
      }
    };
    const onDown = (e: MouseEvent) => {
      const t = e.target as HTMLElement | null;
      if (!menuRef.current) return;
      if (t && menuRef.current.contains(t)) return;
      if (t?.closest('.dialog-backdrop')) return;
      if (t?.closest('[data-testid="cs-dialog-backdrop"]')) return;
      onClose();
    };
    document.addEventListener('keydown', onKey);
    document.addEventListener('mousedown', onDown, true);
    return () => {
      document.removeEventListener('keydown', onKey);
      document.removeEventListener('mousedown', onDown, true);
    };
  }, [onClose, dialogOpen]);

  // The Format dialog side-channels the title change because title
  // text lives on `ChartModel`, not `ChartFormat`. Catch it and apply.
  useEffect(() => {
    const onTitle = (e: Event) => {
      const ce = e as CustomEvent<{ id: string; title?: string }>;
      if (ce.detail.id !== chartId) return;
      update(chartId, { title: ce.detail.title });
    };
    document.addEventListener('casual-chart-title-changed', onTitle);
    return () => document.removeEventListener('casual-chart-title-changed', onTitle);
  }, [chartId, update]);

  if (!chart) return null;

  // Anchor the menu inside the viewport — flip when too close to the
  // right/bottom edges. Excel does the same.
  const W = 200;
  const H = 200; // conservative; menu auto-sizes anyway
  const left = Math.min(x, window.innerWidth - W - 8);
  const top = Math.min(y, window.innerHeight - H - 8);

  if (renaming !== null) {
    return (
      <div
        ref={menuRef}
        className="chart-context-menu chart-context-menu--renaming"
        data-testid="chart-context-menu"
        style={{ position: 'fixed', left, top, pointerEvents: 'auto' }}
      >
        <input
          autoFocus
          className="chart-context-menu__rename-input"
          data-testid="chart-context-rename-input"
          value={renaming}
          onChange={(e) => setRenaming(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === 'Enter') {
              const v = renaming.trim();
              if (v) update(chartId, { title: v });
              onClose();
            }
            if (e.key === 'Escape') onClose();
          }}
          onBlur={() => {
            const v = renaming.trim();
            if (v && v !== chart.title) update(chartId, { title: v });
            onClose();
          }}
        />
      </div>
    );
  }

  return (
    <>
      <div
        ref={menuRef}
        className="chart-context-menu"
        data-testid="chart-context-menu"
        style={{ position: 'fixed', left, top, pointerEvents: 'auto' }}
        role="menu"
      >
        <button
          type="button"
          className="chart-context-menu__item"
          data-testid="chart-context-change-type"
          onClick={() => setShowChangeType(true)}
        >
          <Icon name="bar_chart" />
          <span>Change chart type…</span>
        </button>
        <button
          type="button"
          className="chart-context-menu__item"
          data-testid="chart-context-format"
          onClick={() => setShowFormat(true)}
        >
          <Icon name="tune" />
          <span>Format chart…</span>
        </button>
        <button
          type="button"
          className="chart-context-menu__item"
          data-testid="chart-context-rename"
          onClick={() => setRenaming(chart.title ?? 'Chart')}
        >
          <Icon name="edit" />
          <span>Rename</span>
        </button>
        <div className="chart-context-menu__sep" />
        <button
          type="button"
          className="chart-context-menu__item chart-context-menu__item--danger"
          data-testid="chart-context-delete"
          onClick={() => {
            remove(chartId);
            onClose();
          }}
        >
          <Icon name="delete" />
          <span>Delete chart</span>
        </button>
      </div>

      {showChangeType && api && (
        <InsertChartDialog
          api={api}
          defaultSourceA1={rangeToA1(chart.source)}
          initialType={chart.type}
          title="Change chart type"
          confirmLabel="Apply"
          onCancel={() => {
            setShowChangeType(false);
            onClose();
          }}
          onConfirm={({ source, type }) => {
            update(chartId, { source, type });
            setShowChangeType(false);
            onClose();
          }}
        />
      )}

      {showFormat && (
        <FormatChartDialog
          model={chart}
          seriesNames={readSeriesNames(api, chart)}
          onCancel={() => {
            setShowFormat(false);
            onClose();
          }}
          onConfirm={(next) => {
            update(chartId, { format: next });
            setShowFormat(false);
            onClose();
          }}
        />
      )}
    </>
  );
}

function rangeToA1(r: {
  startRow: number;
  endRow: number;
  startColumn: number;
  endColumn: number;
}): string {
  const a = colA1(r.startColumn) + (r.startRow + 1);
  if (r.startRow === r.endRow && r.startColumn === r.endColumn) return a;
  return `${a}:${colA1(r.endColumn)}${r.endRow + 1}`;
}

function colA1(c: number): string {
  let n = c + 1;
  let s = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function readSeriesNames(
  api: FUniver | null,
  chart: { sheetId: string; source: { startRow: number; startColumn: number; endColumn: number } },
): string[] {
  if (!api) return [];
  try {
    const wb = api.getActiveWorkbook();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const sheets = wb?.getSheets() as any[];
    const ws = sheets?.find((s) => s.getSheetId?.() === chart.sheetId);
    if (!ws) return [];
    const { startRow, startColumn, endColumn } = chart.source;
    const names: string[] = [];
    for (let c = 1; c <= endColumn - startColumn; c += 1) {
      const v = ws.getRange(startRow, startColumn + c).getValue();
      names.push(v == null ? `Series ${c}` : String(v));
    }
    return names;
  } catch (err) {
    console.debug('[chart-format] failed to read series names', err);
    return [];
  }
}
