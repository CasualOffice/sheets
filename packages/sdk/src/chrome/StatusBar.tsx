/**
 * StatusBar — built-in status bar for `<CasualSheets chrome>`.
 *
 * Self-contained: reads the active selection through `CasualSheetsAPI` and shows
 * Excel-style aggregates over it (Average / Count / Numerical Count / Min / Max
 * / Sum). Drives the editor only through the facade.
 *
 * Count = non-empty cells (any type); Numerical Count = numeric cells; the
 * numeric aggregates (Average/Min/Max/Sum) run over the numeric cells only —
 * matching Excel's status-bar semantics.
 *
 * A zoom control is deferred to a follow-up batch: the SDK's current eager
 * plugin set doesn't register `SheetsZoomRenderController`, so the zoom
 * command/operation throws ("Expect 1 dependency item(s)… get 0") in this
 * mount. Wiring zoom needs that registration sorted first.
 */

import { useEffect, useState, type CSSProperties } from 'react';
import { ICommandService } from '@univerjs/core';
import type { CasualSheetsAPI } from '../sheets/api';
// (zoom reader/control deferred — see file header)

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

export interface StatusBarProps {
  api: CasualSheetsAPI | null;
}

export function StatusBar({ api }: StatusBarProps) {
  const [stats, setStats] = useState<Stats | null>(null);

  useEffect(() => {
    if (!api) return;
    const refresh = () => setStats(readStats(api));
    refresh();
    const injector = (api.univer as unknown as { _injector?: { get(t: unknown): unknown } })
      ._injector;
    const cmd = injector?.get(ICommandService) as
      | { onCommandExecuted: (cb: () => void) => { dispose: () => void } }
      | undefined;
    const sub = cmd?.onCommandExecuted(() => refresh());
    return () => sub?.dispose();
  }, [api]);

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
    </div>
  );
}
