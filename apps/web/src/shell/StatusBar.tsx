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
import { useUniverAPI } from '../use-univer';
import { useActiveCellState } from '../hooks/useActiveCellState';
import { redo, undo } from './home-tab-actions';
import { setZoom } from './tab-actions';
import { Icon } from './Icon';
import { Tooltip } from './Tooltip';
import { CollabIndicator } from './CollabIndicator';
import { STAT_LABELS, useStatPrefs, type StatKey } from './use-statbar-prefs';

const NUM = new Intl.NumberFormat(undefined, { maximumFractionDigits: 4 });
const ZOOM_STEPS = [25, 50, 75, 100, 125, 150, 200, 300, 400];

/**
 * Status bar — its own strip below the sheet tabs (Excel/Sheets pattern).
 * Carries the non-tab info that used to crowd the sheet-tabs row: selection
 * stats (Average / Count / Sum / …), collab state, undo/redo, and zoom.
 * Splitting it out means the tabs can grow/scroll without ever colliding with
 * the stats. Reuses the existing strip control styles.
 */
export function StatusBar() {
  const api = useUniverAPI();
  const { stats } = useActiveCellState();

  const [zoomPct, setZoomPct] = useState(100);
  useEffect(() => {
    if (!api) return;
    const d = api.addEvent(api.Event.CommandExecuted, (e) => {
      const info = e as { id?: string; params?: { zoomRatio?: number } };
      if (
        info.id === 'sheet.command.set-zoom-ratio' &&
        typeof info.params?.zoomRatio === 'number'
      ) {
        setZoomPct(Math.round(info.params.zoomRatio * 100));
      }
    });
    return () => d.dispose();
  }, [api]);

  const applyZoom = (pct: number) => {
    if (!api) return;
    const clamped = Math.max(25, Math.min(400, pct));
    setZoomPct(clamped);
    setZoom(api, clamped / 100);
  };

  return (
    <div className="statusbar" data-testid="statusbar" role="status">
      {stats && stats.cellCount > 0 ? (
        <Stats stats={stats} />
      ) : (
        <span className="statusbar__ready" />
      )}

      <div className="statusbar__right">
        <CollabIndicator />
        <span className="statusbar__sep" aria-hidden="true" />
        <Tooltip label="Undo (Ctrl+Z)" side="top">
          <button
            type="button"
            className="statusbar__action btn btn--icon"
            data-testid="qat-undo"
            aria-label="Undo (Ctrl+Z)"
            disabled={!api}
            onClick={() => api && undo(api)}
          >
            <Icon name="undo" size="sm" />
          </button>
        </Tooltip>
        <Tooltip label="Redo (Ctrl+Y)" side="top">
          <button
            type="button"
            className="statusbar__action btn btn--icon"
            data-testid="qat-redo"
            aria-label="Redo (Ctrl+Y)"
            disabled={!api}
            onClick={() => api && redo(api)}
          >
            <Icon name="redo" size="sm" />
          </button>
        </Tooltip>

        <span className="statusbar__sep" aria-hidden="true" />

        <Tooltip label="Zoom out" side="top">
          <button
            type="button"
            className="statusbar__action btn btn--icon"
            data-testid="statusbar-zoom-out"
            aria-label="Zoom out"
            onClick={() => applyZoom([...ZOOM_STEPS].reverse().find((s) => s < zoomPct) ?? 25)}
          >
            <Icon name="zoom_out" size="sm" />
          </button>
        </Tooltip>
        <input
          type="range"
          min={25}
          max={400}
          step={5}
          value={zoomPct}
          data-testid="statusbar-zoom-slider"
          aria-label="Zoom slider"
          className="statusbar__zoom-slider"
          onChange={(e) => applyZoom(Number(e.target.value))}
        />
        <Tooltip label="Zoom in" side="top">
          <button
            type="button"
            className="statusbar__action btn btn--icon"
            data-testid="statusbar-zoom-in"
            aria-label="Zoom in"
            onClick={() => applyZoom(ZOOM_STEPS.find((s) => s > zoomPct) ?? 400)}
          >
            <Icon name="zoom_in" size="sm" />
          </button>
        </Tooltip>
        <Tooltip label="Reset to 100%" side="top">
          <button
            type="button"
            className="statusbar__zoom-label"
            data-testid="statusbar-zoom-label"
            aria-label="Reset zoom to 100%"
            onClick={() => applyZoom(100)}
          >
            {zoomPct}%
          </button>
        </Tooltip>
      </div>
    </div>
  );
}

/**
 * Selection stats with Excel-style right-click customisation (Average /
 * Count / Numerical Count / Min / Max / Sum). Prefs persist via useStatPrefs.
 */
function Stats({ stats }: { stats: NonNullable<ReturnType<typeof useActiveCellState>['stats']> }) {
  const { prefs, toggle } = useStatPrefs();
  const [menuOpen, setMenuOpen] = useState(false);
  const stripRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!menuOpen) return;
    const onDown = (e: PointerEvent) => {
      if (!stripRef.current?.contains(e.target as Node)) setMenuOpen(false);
    };
    document.addEventListener('pointerdown', onDown, true);
    return () => document.removeEventListener('pointerdown', onDown, true);
  }, [menuOpen]);

  const items: Array<{ key: StatKey; node: React.ReactNode }> = [];
  if (prefs.avg && stats.avg !== null) {
    items.push({
      key: 'avg',
      node: <span data-testid="stat-avg">Average: {NUM.format(stats.avg)}</span>,
    });
  }
  if (prefs.count) {
    items.push({
      key: 'count',
      node: <span data-testid="stat-count">Count: {stats.cellCount}</span>,
    });
  }
  if (prefs.numCount && stats.count !== stats.cellCount) {
    items.push({
      key: 'numCount',
      node: <span data-testid="stat-num-count">Numerical Count: {stats.count}</span>,
    });
  }
  if (prefs.min && stats.min !== null) {
    items.push({
      key: 'min',
      node: <span data-testid="stat-min">Min: {NUM.format(stats.min)}</span>,
    });
  }
  if (prefs.max && stats.max !== null) {
    items.push({
      key: 'max',
      node: <span data-testid="stat-max">Max: {NUM.format(stats.max)}</span>,
    });
  }
  if (prefs.sum) {
    items.push({
      key: 'sum',
      node: <span data-testid="stat-sum">Sum: {NUM.format(stats.sum)}</span>,
    });
  }

  return (
    <div
      ref={stripRef}
      className="statusbar__stats"
      data-testid="sheet-tabs-stats"
      onContextMenu={(e) => {
        e.preventDefault();
        setMenuOpen((v) => !v);
      }}
      title="Right-click to choose which stats appear"
    >
      {items.map((it) => (
        <span key={it.key}>{it.node}</span>
      ))}
      {menuOpen && (
        <div className="statbar-customise" role="menu" data-testid="statbar-customise">
          <div className="statbar-customise__heading">Customise Status Bar</div>
          {(Object.keys(STAT_LABELS) as StatKey[]).map((key) => (
            <label
              key={key}
              className="statbar-customise__item"
              data-testid={`statbar-customise-${key}`}
            >
              <input type="checkbox" checked={prefs[key]} onChange={() => toggle(key)} />
              <span>{STAT_LABELS[key]}</span>
            </label>
          ))}
        </div>
      )}
    </div>
  );
}
