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
 * InsertChartDialog — the SDK chrome's built-in Insert Chart modal.
 *
 * Reads the active A1 selection off the FUniver facade, lets the user pick a
 * chart type (column / bar / line / pie) and confirm the source range, then
 * inserts a chart over the grid.
 *
 * Facade grounding / capability note
 * ----------------------------------
 * Univer 0.25 (the version pinned in this SDK's peerDependencies) ships NO
 * `@univerjs/*chart*` package — there is no `FUniver.newChart()` builder and no
 * `FRange.insertChart()` the way `@univerjs/sheets-data-validation` gives us
 * `newDataValidation()` / `setDataValidation()`. (Verified: no chart package
 * under node_modules/@univerjs, and none listed in the SDK peerDependencies.)
 *
 * The closest REAL, installed facade capability for placing a positioned visual
 * artifact over the grid is the float-DOM API contributed by
 * `@univerjs/sheets-drawing-ui/facade` — the same mechanism Univer's own chart
 * and image plugins use to anchor content over cells:
 *   - `univerAPI.registerComponent(key, Component)`  (@univerjs/ui f-univer.d.ts L390)
 *   - `FWorksheet.addFloatDomToRange(range, { componentKey, data }, domLayout, id)`
 *     (@univerjs/sheets-drawing-ui/lib/types/facade/f-worksheet.d.ts L305; the
 *      `declare module '@univerjs/sheets/facade' { interface FWorksheet … }`
 *      augmentation at L428 is what puts the method on the live sheet)
 *
 * So on Apply we: register a small self-contained SVG chart-preview component
 * (once), read the selected range's numeric values off the SDK snapshot, and
 * insert a float DOM over the selection rendering that preview for the chosen
 * type. This is a genuine facade insertion (real, disposable, positioned),
 * reads the real selection data, and closes.
 *
 * LIMITATION (recorded): without a chart engine in the SDK's dep tree, the
 * inserted artifact is a static, computed SVG preview of the selected values —
 * not a live, re-calculating, fully-styled chart series. When a host installs a
 * real Univer chart plugin it can OVERRIDE this dialog by registering its own
 * `insert-chart` extension (extensions.dialogs['insert-chart']); the built-in is
 * the always-available default.
 *
 * Mounted by `<DialogHost>` when `openDialog('insert-chart')` is called and no
 * host override is registered.
 */

import { useMemo, useState, type CSSProperties, type ReactNode } from 'react';
// Side-effect import: installs `FWorksheet.addFloatDomToRange` (+ the other
// float-dom methods) on the facade prototype AND augments
// `@univerjs/sheets/facade`'s FWorksheet at type-check time. Without it, the
// method is undefined at runtime. Mirrors DataValidationDialog's
// `@univerjs/sheets-data-validation/facade` side-effect import.
import '@univerjs/sheets-drawing-ui/facade';
import type { DialogComponentProps } from './extensions';
import type { CasualSheetsAPI } from '../sheets/api';
import { Dialog } from './Dialog';
import {
  DIALOG_BTN_PRIMARY_STYLE,
  DIALOG_BTN_SECONDARY_STYLE,
  DIALOG_FIELD_STYLE,
  DIALOG_INPUT_STYLE,
  DIALOG_LABEL_STYLE,
} from './dialog-styles';

/** Chart families the dialog can insert. */
type ChartType = 'column' | 'bar' | 'line' | 'pie';

const CHART_TYPE_OPTIONS: Array<{ value: ChartType; label: string }> = [
  { value: 'column', label: 'Column' },
  { value: 'bar', label: 'Bar' },
  { value: 'line', label: 'Line' },
  { value: 'pie', label: 'Pie' },
];

/** Stable component key we register the preview under (once per FUniver). */
const CHART_COMPONENT_KEY = 'cs-builtin-chart-preview';

/** Palette for the preview series (kept small + colour-blind-friendly-ish). */
const PALETTE = ['#0e7490', '#f59e0b', '#10b981', '#ef4444', '#6366f1', '#ec4899'];

interface DialogState {
  chartType: ChartType;
  /** A1 range the chart reads from (seeded from the live selection). */
  rangeA1: string;
}

/** The active FRange, or null when there is no selection. */
function activeRange(api: CasualSheetsAPI) {
  return api.univer.getActiveWorkbook()?.getActiveSheet()?.getActiveRange() ?? null;
}

/** The active FWorksheet, or null. */
function activeSheet(api: CasualSheetsAPI) {
  return api.univer.getActiveWorkbook()?.getActiveSheet() ?? null;
}

/**
 * Pull the numeric values out of the selected range off the SDK snapshot (the
 * same read path FormatCellsDialog uses to seed itself). Flattened + coerced;
 * non-numeric cells drop out. Returns [] when nothing usable is present.
 */
function readRangeNumbers(api: CasualSheetsAPI): number[] {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const snap = api.getSnapshot() as any;
  const sel = api.getSelection();
  if (!snap || !sel) return [];
  const sheet = snap.sheets?.[sel.sheetId];
  if (!sheet) return [];
  const { startRow, endRow, startColumn, endColumn } = sel.range;
  const out: number[] = [];
  for (let r = startRow; r <= endRow; r++) {
    for (let c = startColumn; c <= endColumn; c++) {
      const raw = sheet.cellData?.[r]?.[c]?.v;
      const n = typeof raw === 'number' ? raw : Number(raw);
      if (Number.isFinite(n)) out.push(n);
    }
  }
  return out;
}

/**
 * Register the SVG chart-preview component on the FUniver facade if it hasn't
 * been registered yet. Idempotent per FUniver instance: we stash a flag on the
 * facade object so re-opening the dialog doesn't re-register (registerComponent
 * returns a disposable, but we intentionally keep the registration alive for the
 * lifetime of the editor so previously-inserted charts keep rendering).
 */
function ensureChartComponent(api: CasualSheetsAPI): void {
  const uni = api.univer as unknown as {
    registerComponent?: (key: string, comp: unknown) => { dispose(): void };
    __csChartRegistered?: boolean;
  };
  if (uni.__csChartRegistered || typeof uni.registerComponent !== 'function') return;
  uni.registerComponent(CHART_COMPONENT_KEY, ChartPreview);
  uni.__csChartRegistered = true;
}

/**
 * Insert the chart over the current selection via the drawing-ui float-DOM
 * facade. Returns false when there's no selection or the facade method is
 * missing (host didn't register the drawing-ui plugin). Reads the range's
 * numbers and hands them to the preview component as `data`.
 */
function insertChart(api: CasualSheetsAPI, s: DialogState): boolean {
  const range = activeRange(api);
  const sheet = activeSheet(api);
  if (!range || !sheet) return false;

  ensureChartComponent(api);

  const values = readRangeNumbers(api);

  // `addFloatDomToRange` is contributed by @univerjs/sheets-drawing-ui/facade
  // (f-worksheet.d.ts L305). Typed loosely at the call site so the SDK doesn't
  // hard-depend on the ambient FWorksheet augmentation here.
  const domSheet = sheet as unknown as {
    addFloatDomToRange?: (
      range: unknown,
      layer: { componentKey: string; data?: unknown; allowTransform?: boolean },
      domLayout: {
        width?: number;
        height?: number;
        marginX?: number | string;
        marginY?: number | string;
      },
      id?: string,
    ) => { id: string; dispose: () => void } | null;
  };
  if (typeof domSheet.addFloatDomToRange !== 'function') return false;

  const result = domSheet.addFloatDomToRange(
    range,
    {
      componentKey: CHART_COMPONENT_KEY,
      allowTransform: true,
      data: { type: s.chartType, values, rangeA1: s.rangeA1 },
    },
    { width: 360, height: 240, marginX: 0, marginY: 0 },
    `cs-chart-${Date.now()}`,
  );
  return result != null;
}

/**
 * Self-contained SVG chart preview rendered inside the float DOM. Univer passes
 * the `layer.data` object through as the component's `data` prop. Deliberately
 * dependency-free (no echarts) so the SDK bundle stays lean and the built-in
 * works without any host chart engine.
 */
function ChartPreview({
  data,
}: {
  data?: { type?: ChartType; values?: number[]; rangeA1?: string };
}) {
  const type = data?.type ?? 'column';
  const values = (data?.values ?? []).slice(0, 24);
  const label = data?.rangeA1 ?? '';

  const wrapStyle: CSSProperties = {
    width: '100%',
    height: '100%',
    boxSizing: 'border-box',
    background: '#ffffff',
    border: '1px solid #cdd3db',
    borderRadius: 6,
    padding: 8,
    display: 'flex',
    flexDirection: 'column',
    font: '12px system-ui, sans-serif',
    color: '#201f1e',
    overflow: 'hidden',
  };
  const titleStyle: CSSProperties = {
    fontWeight: 600,
    marginBottom: 4,
    whiteSpace: 'nowrap',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
  };

  const W = 340;
  const H = 176;
  const max = values.length ? Math.max(...values, 0) : 0;
  const min = values.length ? Math.min(...values, 0) : 0;
  const span = max - min || 1;

  let body: ReactNode;
  if (values.length === 0) {
    body = (
      <text x={W / 2} y={H / 2} textAnchor="middle" fill="#605e5c">
        No numeric data in range
      </text>
    );
  } else if (type === 'pie') {
    const total = values.reduce((a, v) => a + Math.abs(v), 0) || 1;
    const cx = W / 2;
    const cy = H / 2;
    const rad = Math.min(W, H) / 2 - 6;
    let angle = -Math.PI / 2;
    body = (
      <>
        {values.map((v, i) => {
          const frac = Math.abs(v) / total;
          const next = angle + frac * Math.PI * 2;
          const x1 = cx + rad * Math.cos(angle);
          const y1 = cy + rad * Math.sin(angle);
          const x2 = cx + rad * Math.cos(next);
          const y2 = cy + rad * Math.sin(next);
          const large = frac > 0.5 ? 1 : 0;
          const d = `M ${cx} ${cy} L ${x1} ${y1} A ${rad} ${rad} 0 ${large} 1 ${x2} ${y2} Z`;
          angle = next;
          return <path key={i} d={d} fill={PALETTE[i % PALETTE.length]} />;
        })}
      </>
    );
  } else if (type === 'line') {
    const step = values.length > 1 ? W / (values.length - 1) : W;
    const pts = values.map((v, i) => `${i * step},${H - ((v - min) / span) * H}`).join(' ');
    body = <polyline points={pts} fill="none" stroke={PALETTE[0]} strokeWidth={2} />;
  } else {
    // column (vertical) or bar (horizontal)
    const horizontal = type === 'bar';
    const n = values.length;
    const gap = 4;
    body = (
      <>
        {values.map((v, i) => {
          const fill = PALETTE[i % PALETTE.length];
          if (horizontal) {
            const bh = H / n - gap;
            const bw = ((v - min) / span) * W;
            return (
              <rect
                key={i}
                x={0}
                y={i * (bh + gap)}
                width={Math.max(bw, 1)}
                height={bh}
                fill={fill}
              />
            );
          }
          const bw = W / n - gap;
          const bh = ((v - min) / span) * H;
          return (
            <rect
              key={i}
              x={i * (bw + gap)}
              y={H - bh}
              width={bw}
              height={Math.max(bh, 1)}
              fill={fill}
            />
          );
        })}
      </>
    );
  }

  return (
    <div style={wrapStyle}>
      {label && (
        <div style={titleStyle}>{`${type[0].toUpperCase()}${type.slice(1)} chart · ${label}`}</div>
      )}
      <svg viewBox={`0 0 ${W} ${H}`} width="100%" height="100%" preserveAspectRatio="xMidYMid meet">
        {body}
      </svg>
    </div>
  );
}

const RANGE_NOTE_STYLE: CSSProperties = {
  fontSize: 12,
  color: 'var(--cs-chrome-muted, #605e5c)',
  marginBottom: 12,
};

const TYPE_ROW_STYLE: CSSProperties = {
  display: 'grid',
  gridTemplateColumns: 'repeat(4, 1fr)',
  gap: 8,
  marginBottom: 12,
};

function typeCardStyle(selected: boolean): CSSProperties {
  return {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    gap: 6,
    padding: '10px 4px',
    border: `1px solid ${selected ? 'var(--cs-chrome-active-fg, #0e7490)' : 'var(--cs-chrome-border, #cdd3db)'}`,
    borderRadius: 8,
    background: selected
      ? 'var(--cs-chrome-active-bg, #ecfeff)'
      : 'var(--cs-chrome-input-bg, #fff)',
    color: selected ? 'var(--cs-chrome-active-fg, #0e7490)' : 'var(--cs-chrome-fg, #201f1e)',
    font: 'inherit',
    fontSize: 12,
    cursor: 'pointer',
  };
}

/** Tiny inline SVG glyph per chart type — no icon-font dependency. */
function TypeGlyph({ type }: { type: ChartType }) {
  const c = 'currentColor';
  switch (type) {
    case 'column':
      return (
        <svg width={22} height={22} viewBox="0 0 22 22" fill={c} aria-hidden="true">
          <rect x={2} y={10} width={4} height={10} />
          <rect x={9} y={5} width={4} height={15} />
          <rect x={16} y={2} width={4} height={18} />
        </svg>
      );
    case 'bar':
      return (
        <svg width={22} height={22} viewBox="0 0 22 22" fill={c} aria-hidden="true">
          <rect x={2} y={2} width={10} height={4} />
          <rect x={2} y={9} width={16} height={4} />
          <rect x={2} y={16} width={7} height={4} />
        </svg>
      );
    case 'line':
      return (
        <svg
          width={22}
          height={22}
          viewBox="0 0 22 22"
          fill="none"
          stroke={c}
          strokeWidth={2}
          aria-hidden="true"
        >
          <polyline points="2,16 8,9 13,13 20,3" />
        </svg>
      );
    case 'pie':
      return (
        <svg width={22} height={22} viewBox="0 0 22 22" fill={c} aria-hidden="true">
          <path d="M11 11 L11 2 A9 9 0 0 1 20 11 Z" />
          <circle cx={11} cy={11} r={9} fill="none" stroke={c} strokeWidth={1.5} />
        </svg>
      );
  }
}

export function InsertChartDialog({ api, onClose }: DialogComponentProps) {
  // Read the selection once for the header hint + as the seed source range.
  // `getA1Notation` off the live FRange (verified in @univerjs/sheets/facade
  // f-range.d.ts) gives the user-facing A1 label, e.g. "A1:B5".
  const rangeLabel = useMemo(() => {
    const fRange = activeRange(api) as unknown as { getA1Notation?: () => string } | null;
    return fRange?.getA1Notation?.() ?? null;
  }, [api]);

  const [state, setState] = useState<DialogState>(() => ({
    chartType: 'column',
    rangeA1: rangeLabel ?? '',
  }));

  const hasSelection = activeRange(api) !== null;

  const update = <K extends keyof DialogState>(key: K, value: DialogState[K]) =>
    setState((prev) => ({ ...prev, [key]: value }));

  const apply = () => {
    if (insertChart(api, state)) onClose();
  };

  return (
    <Dialog
      title="Insert chart"
      onClose={onClose}
      width={440}
      data-testid="cs-insert-chart-dialog"
      footer={
        <>
          <button type="button" style={DIALOG_BTN_SECONDARY_STYLE} onClick={onClose}>
            Cancel
          </button>
          <button
            type="button"
            style={DIALOG_BTN_PRIMARY_STYLE}
            data-testid="cs-insert-chart-apply"
            disabled={!hasSelection}
            onClick={apply}
          >
            Insert
          </button>
        </>
      }
    >
      {hasSelection ? (
        <div style={RANGE_NOTE_STYLE} data-testid="cs-insert-chart-range">
          Chart from <strong>{rangeLabel ?? 'the current selection'}</strong>
        </div>
      ) : (
        <div style={RANGE_NOTE_STYLE} data-testid="cs-insert-chart-no-selection">
          Select the data range first, then reopen this dialog.
        </div>
      )}

      <div style={DIALOG_LABEL_STYLE}>Chart type</div>
      <div style={TYPE_ROW_STYLE} role="radiogroup" aria-label="Chart type">
        {CHART_TYPE_OPTIONS.map((opt) => (
          <button
            key={opt.value}
            type="button"
            role="radio"
            aria-checked={state.chartType === opt.value}
            data-testid={`cs-insert-chart-type-${opt.value}`}
            style={typeCardStyle(state.chartType === opt.value)}
            onClick={() => update('chartType', opt.value)}
          >
            <TypeGlyph type={opt.value} />
            <span>{opt.label}</span>
          </button>
        ))}
      </div>

      <label style={DIALOG_FIELD_STYLE}>
        <span style={DIALOG_LABEL_STYLE}>Data range</span>
        <input
          style={DIALOG_INPUT_STYLE}
          data-testid="cs-insert-chart-range-input"
          value={state.rangeA1}
          placeholder="A1:B5"
          onChange={(e) => update('rangeA1', e.target.value)}
        />
      </label>
      <div style={RANGE_NOTE_STYLE}>
        The chart is placed over the current selection. Numeric cells in the range become the
        series.
      </div>
    </Dialog>
  );
}
