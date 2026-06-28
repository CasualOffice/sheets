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
 * StatusBar — built-in status bar for `<CasualSheets chrome>`.
 *
 * Self-contained: reads the active selection through `CasualSheetsAPI` and shows
 * Excel-style aggregates over it (Average / Count / Numerical Count / Min / Max
 * / Sum), plus a zoom control on the right (− / level / +, click the level to
 * reset to 100%). Drives the editor only through the facade + commands.
 *
 * Count = non-empty cells (any type); Numerical Count = numeric cells; the
 * numeric aggregates (Average/Min/Max/Sum) run over the numeric cells only —
 * matching Excel's status-bar semantics.
 *
 * Zoom dispatches the OPERATION `sheet.operation.set-zoom-ratio` (not the
 * `set-zoom-ratio` / `change-zoom-ratio` commands, which bail when Univer's
 * always-present formula-bar editor unit reports visible). The operation's
 * render controller registers in sheets-ui's `onRendered` lifecycle — i.e. a
 * frame or two after `onReady` — so the buttons just no-op until the grid has
 * painted; in practice the user can't click that fast.
 */

import { useEffect, useState, type CSSProperties } from 'react';
import { ICommandService } from '@univerjs/core';
import type { CasualSheetsAPI } from '../sheets/api';

interface Stats {
  /** Non-empty cells (any type). */
  count: number;
  /** Numeric cells. */
  numCount: number;
  sum: number;
  avg: number;
  min: number;
  max: number;
}

function readStats(api: CasualSheetsAPI): Stats | null {
  const sel = api.getSelection();
  const sheet = api.univer.getActiveWorkbook()?.getActiveSheet();
  if (!sel || !sheet) return null;
  const { startRow, startColumn, endRow, endColumn } = sel.range;
  // Single cell → nothing to aggregate (matches Excel).
  if (startRow === endRow && startColumn === endColumn) return null;
  const values = sheet.getRange(sel.range).getValues?.() as unknown[][] | undefined;
  if (!values) return null;
  let count = 0;
  let numCount = 0;
  let sum = 0;
  let min = Infinity;
  let max = -Infinity;
  for (const row of values) {
    for (const v of row) {
      if (v == null || v === '') continue;
      count += 1;
      if (typeof v === 'number' && Number.isFinite(v)) {
        numCount += 1;
        sum += v;
        if (v < min) min = v;
        if (v > max) max = v;
      }
    }
  }
  if (count === 0) return null;
  return { count, numCount, sum, avg: numCount ? sum / numCount : 0, min, max };
}

/** Current zoom as a fraction (1 = 100%). Reads the live worksheet config via
 *  the facade's `getSheet()` escape hatch; falls back to 1. */
function readZoom(api: CasualSheetsAPI): number {
  const sheet = api.univer.getActiveWorkbook()?.getActiveSheet();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const ws = (sheet as any)?.getSheet?.();
  const z = ws?.getConfig?.().zoomRatio;
  return typeof z === 'number' && z > 0 ? z : 1;
}

// Trim float noise without locking to a fixed precision.
function fmt(n: number): string {
  return Number(n.toFixed(10)).toLocaleString();
}

const BAR_STYLE: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'flex-end',
  gap: 16,
  height: 24,
  padding: '0 12px',
  borderTop: '1px solid var(--cs-chrome-border, rgba(0,0,0,0.12))',
  background: 'var(--cs-chrome-bg, #f8f9fa)',
  color: 'var(--cs-chrome-muted, #4b5563)',
  flex: '0 0 auto',
  font: 'inherit',
  fontSize: 12,
  userSelect: 'none',
};

const ZOOM_GROUP_STYLE: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 4,
};

const ZOOM_BTN_STYLE: CSSProperties = {
  width: 20,
  height: 18,
  border: 'none',
  borderRadius: 4,
  background: 'transparent',
  color: 'var(--cs-chrome-fg, #201f1e)',
  cursor: 'pointer',
  fontSize: 14,
  lineHeight: 1,
  padding: 0,
};

const ZOOM_LEVEL_STYLE: CSSProperties = {
  minWidth: 40,
  height: 18,
  border: 'none',
  borderRadius: 4,
  background: 'transparent',
  color: 'var(--cs-chrome-muted, #4b5563)',
  cursor: 'pointer',
  font: 'inherit',
  fontSize: 12,
  textAlign: 'center',
};

export interface StatusBarProps {
  api: CasualSheetsAPI | null;
}

export function StatusBar({ api }: StatusBarProps) {
  const [stats, setStats] = useState<Stats | null>(null);
  const [zoom, setZoom] = useState(1);

  useEffect(() => {
    if (!api) return;
    const refresh = () => {
      setStats(readStats(api));
      setZoom(readZoom(api));
    };
    refresh();
    const injector = (api.univer as unknown as { _injector?: { get(t: unknown): unknown } })
      ._injector;
    const cmd = injector?.get(ICommandService) as
      | { onCommandExecuted: (cb: () => void) => { dispose: () => void } }
      | undefined;
    const sub = cmd?.onCommandExecuted(() => refresh());
    return () => sub?.dispose();
  }, [api]);

  // Drive the zoom OPERATION directly — see file header for why not the commands.
  const setZoomRatio = (ratio: number) => {
    const wb = api?.univer.getActiveWorkbook();
    const sheet = wb?.getActiveSheet();
    if (!wb || !sheet) return;
    const clamped = Math.min(4, Math.max(0.1, ratio));
    void api?.executeCommand('sheet.operation.set-zoom-ratio', {
      unitId: wb.getId(),
      subUnitId: sheet.getSheetId(),
      zoomRatio: clamped,
    });
    // Optimistic local update — the onCommandExecuted refresh will reconcile.
    setZoom(clamped);
  };
  const zoomBy = (delta: number) => setZoomRatio(Math.round((zoom + delta) * 100) / 100);
  const resetZoom = () => setZoomRatio(1);

  return (
    <div style={BAR_STYLE} data-testid="casual-sheets-status-bar">
      {stats && (
        <>
          {stats.numCount > 0 && <span data-stat="average">Average: {fmt(stats.avg)}</span>}
          <span data-stat="count">Count: {stats.count}</span>
          {stats.numCount > 0 && (
            <>
              <span data-stat="num-count">Numerical Count: {stats.numCount}</span>
              <span data-stat="min">Min: {fmt(stats.min)}</span>
              <span data-stat="max">Max: {fmt(stats.max)}</span>
              <span data-stat="sum">Sum: {fmt(stats.sum)}</span>
            </>
          )}
        </>
      )}
      <span style={ZOOM_GROUP_STYLE} data-testid="cs-zoom">
        <button
          type="button"
          style={ZOOM_BTN_STYLE}
          aria-label="Zoom out"
          title="Zoom out"
          data-testid="cs-zoom-out"
          disabled={!api}
          onMouseDown={(e) => {
            e.preventDefault();
            zoomBy(-0.1);
          }}
        >
          −
        </button>
        <button
          type="button"
          style={ZOOM_LEVEL_STYLE}
          aria-label="Reset zoom to 100%"
          title="Reset zoom to 100%"
          data-testid="cs-zoom-level"
          disabled={!api}
          onMouseDown={(e) => {
            e.preventDefault();
            resetZoom();
          }}
        >
          {Math.round(zoom * 100)}%
        </button>
        <button
          type="button"
          style={ZOOM_BTN_STYLE}
          aria-label="Zoom in"
          title="Zoom in"
          data-testid="cs-zoom-in"
          disabled={!api}
          onMouseDown={(e) => {
            e.preventDefault();
            zoomBy(0.1);
          }}
        >
          +
        </button>
      </span>
    </div>
  );
}
