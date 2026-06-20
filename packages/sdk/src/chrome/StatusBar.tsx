/**
 * StatusBar — minimal built-in status bar for `<CasualSheets chrome>`.
 *
 * Third chrome slice. Self-contained: reads the active selection through
 * `CasualSheetsAPI` and shows Excel-style aggregates (Average / Count / Sum)
 * over the numeric cells in it. No app context. The richer status bar
 * (configurable stats, min/max, zoom, sheet tabs) lifts behind `chrome="full"`.
 */

import { useEffect, useState, type CSSProperties } from 'react';
import { ICommandService } from '@univerjs/core';
import type { CasualSheetsAPI } from '../sheets/api';

interface Stats {
  count: number;
  sum: number;
  avg: number;
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
  let sum = 0;
  for (const row of values) {
    for (const v of row) {
      if (typeof v === 'number' && Number.isFinite(v)) {
        count += 1;
        sum += v;
      }
    }
  }
  if (count === 0) return null;
  return { count, sum, avg: sum / count };
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
          <span data-stat="average">Average: {fmt(stats.avg)}</span>
          <span data-stat="count">Count: {stats.count}</span>
          <span data-stat="sum">Sum: {fmt(stats.sum)}</span>
        </>
      )}
    </div>
  );
}
