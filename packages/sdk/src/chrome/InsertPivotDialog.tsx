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
 * InsertPivotDialog — the SDK chrome's built-in "Insert pivot table" modal
 * (dialog kind `'insert-pivot'`).
 *
 * Facade note / LIMITATION
 * ------------------------
 * Univer's first-class pivot-table feature lives in the commercial
 * `@univerjs/sheets-pivot(-ui)` packages, which are NOT installed in this SDK
 * (the workspace ships only the OSS `@univerjs/*` set — verified: no `pivot`
 * package under `packages/sdk/node_modules/@univerjs/` or `.pnpm/`, and none in
 * `package.json`). So there is no `univerAPI.newPivotTable()` / FRange pivot
 * facade to call.
 *
 * Rather than stub, this dialog implements a REAL, self-contained group-by
 * summary — the practical core of a pivot — entirely on the installed
 * `@univerjs/sheets` facade:
 *   - reads the active source range's values via `FRange.getValues()`
 *     (f-range.d.ts L386),
 *   - groups rows by a chosen "Rows" column and aggregates a chosen "Values"
 *     column (Sum · Count · Average · Min · Max),
 *   - writes the resulting table to the destination:
 *       • New sheet     → `FWorkbook.create(name, rows, cols)` (f-workbook.d.ts
 *                          L178) then `FWorksheet.getRange(...).setValues(...)`.
 *       • This location  → `FWorksheet.getRange(a1).setValues(...)`
 *                          (f-worksheet.d.ts L279 / f-range.d.ts L953).
 *
 * When the real pivot packages are added later, swap `buildPivot`/`writePivot`
 * for the `newPivotTable()` builder; the dialog's inputs already map 1:1 to a
 * pivot's Rows / Values / aggregation configuration.
 *
 * Mounted by `<DialogHost>` when `openDialog('insert-pivot')` is called and no
 * host override is registered.
 */

import { useMemo, useState, type CSSProperties } from 'react';
import type { IWorkbookData } from '@univerjs/core';
import type { DialogComponentProps } from './extensions';
import type { CasualSheetsAPI } from '../sheets/api';
import type { PivotModel } from '../pivots/types';
import { Dialog } from './Dialog';
import {
  DIALOG_BTN_PRIMARY_STYLE,
  DIALOG_BTN_SECONDARY_STYLE,
  DIALOG_FIELD_STYLE,
  DIALOG_INPUT_STYLE,
  DIALOG_LABEL_STYLE,
} from './dialog-styles';

/** Aggregation functions the summary can apply to the Values column. */
type AggFn = 'sum' | 'count' | 'average' | 'min' | 'max';

/** Where the resulting pivot table is written. */
type Destination = 'newSheet' | 'existing';

const AGG_OPTIONS: Array<{ value: AggFn; label: string }> = [
  { value: 'sum', label: 'Sum' },
  { value: 'count', label: 'Count' },
  { value: 'average', label: 'Average' },
  { value: 'min', label: 'Min' },
  { value: 'max', label: 'Max' },
];

/** A cell value as read off the facade — Univer returns primitives or null. */
type Cell = string | number | boolean | null;

/** The active FRange, or null when there is no selection. */
function activeRange(api: CasualSheetsAPI) {
  return api.univer.getActiveWorkbook()?.getActiveSheet()?.getActiveRange() ?? null;
}

/** Loosely-typed FRange view — only the facade methods this dialog calls. */
interface RangeView {
  getValues: () => Cell[][];
  getA1Notation?: () => string;
  /** Top-left cell coords (0-based) — used to anchor the pivot's source range. */
  getRow?: () => number;
  getColumn?: () => number;
}

/** Loosely-typed FWorksheet view — sheet name + a range handle by A1. */
interface SheetView {
  getSheetName: () => string;
  getSheetId?: () => string;
  getRange: (a1: string) => RangeView & { setValues: (v: Cell[][]) => unknown };
}

/** Read the source grid once. Returns [] when there's no usable selection. */
function readSource(api: CasualSheetsAPI): Cell[][] {
  const range = activeRange(api) as unknown as RangeView | null;
  const grid = range?.getValues?.() ?? [];
  return Array.isArray(grid) ? grid : [];
}

/** A column label for the header dropdowns — "A", "B", or a header-row title. */
function columnLabel(index: number, headerRow: Cell[] | undefined, useHeader: boolean): string {
  if (useHeader) {
    const raw = headerRow?.[index];
    if (raw != null && String(raw).trim().length > 0) return String(raw);
  }
  // Spreadsheet-style A, B, …, Z, AA — relative to the source's first column.
  let n = index;
  let label = '';
  do {
    label = String.fromCharCode(65 + (n % 26)) + label;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return `Column ${label}`;
}

function toNumber(v: Cell): number | null {
  if (typeof v === 'number') return Number.isFinite(v) ? v : null;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (typeof v === 'string' && v.trim() !== '') {
    const n = Number(v);
    return Number.isFinite(n) ? n : null;
  }
  return null;
}

/** Apply the chosen aggregation to a bucket of numeric values. */
function aggregate(fn: AggFn, values: number[], rowCount: number): number {
  if (fn === 'count') return rowCount;
  if (values.length === 0) return 0;
  switch (fn) {
    case 'sum':
      return values.reduce((a, b) => a + b, 0);
    case 'average':
      return values.reduce((a, b) => a + b, 0) / values.length;
    case 'min':
      return Math.min(...values);
    case 'max':
      return Math.max(...values);
  }
}

interface PivotResult {
  /** Full 2D grid (header row + one row per group) ready for setValues. */
  grid: Cell[][];
}

/**
 * Build the group-by summary grid from the source values. Groups the data rows
 * by `groupCol`, aggregates `valueCol` with `fn`, and returns a 2-column table:
 * [group label, aggregated value], preceded by a header row.
 */
function buildPivot(
  source: Cell[][],
  useHeader: boolean,
  groupCol: number,
  valueCol: number,
  fn: AggFn,
): PivotResult | null {
  if (source.length === 0) return null;
  const headerRow = useHeader ? source[0] : undefined;
  const dataRows = useHeader ? source.slice(1) : source;
  if (dataRows.length === 0) return null;

  // Preserve first-seen group order for a stable, predictable output.
  const order: string[] = [];
  const buckets = new Map<string, number[]>();
  const counts = new Map<string, number>();

  for (const row of dataRows) {
    const rawKey = row[groupCol];
    const key = rawKey == null ? '' : String(rawKey);
    if (!buckets.has(key)) {
      buckets.set(key, []);
      counts.set(key, 0);
      order.push(key);
    }
    counts.set(key, (counts.get(key) ?? 0) + 1);
    const num = toNumber(row[valueCol]);
    if (num != null) buckets.get(key)!.push(num);
  }

  const groupHeader = columnLabel(groupCol, headerRow, useHeader);
  const valueHeader = columnLabel(valueCol, headerRow, useHeader);
  const aggLabel = AGG_OPTIONS.find((o) => o.value === fn)?.label ?? 'Sum';

  const grid: Cell[][] = [[groupHeader, `${aggLabel} of ${valueHeader}`]];
  for (const key of order) {
    grid.push([key === '' ? '(blank)' : key, aggregate(fn, buckets.get(key)!, counts.get(key)!)]);
  }
  return { grid };
}

/** Convert a 0-based (row, col) to A1, e.g. (0,0) -> "A1". */
function toA1(row: number, col: number): string {
  let n = col;
  let letters = '';
  do {
    letters = String.fromCharCode(65 + (n % 26)) + letters;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return `${letters}${row + 1}`;
}

/** Resource envelope the pivot models live in — matches PivotFieldsPanel so the
 *  panel edits the same pivot the dialog created. */
const PIVOTS_RESOURCE_NAME = '__casual_sheets_pivots__';

/** Append a freshly-created pivot model to the workbook snapshot resource. */
function persistPivotModel(api: CasualSheetsAPI, model: PivotModel): void {
  const data = api.getContent();
  if (!data) return;
  const resources = data.resources ? [...data.resources] : [];
  const idx = resources.findIndex((r) => r.name === PIVOTS_RESOURCE_NAME);
  let pivots: PivotModel[] = [];
  if (idx >= 0 && resources[idx]?.data) {
    try {
      const parsed = JSON.parse(resources[idx].data) as { v?: number; pivots?: PivotModel[] };
      if (parsed?.v === 1 && Array.isArray(parsed.pivots)) pivots = parsed.pivots;
    } catch {
      pivots = [];
    }
  }
  const merged = pivots.some((p) => p.id === model.id)
    ? pivots.map((p) => (p.id === model.id ? model : p))
    : [...pivots, model];
  const entry = { name: PIVOTS_RESOURCE_NAME, data: JSON.stringify({ v: 1, pivots: merged }) };
  if (idx >= 0) resources[idx] = entry;
  else resources.push(entry);
  const nextData: IWorkbookData = { ...data, resources };
  api.setContent(nextData);
}

interface DialogState {
  useHeader: boolean;
  groupCol: number;
  valueCol: number;
  aggFn: AggFn;
  destination: Destination;
  /** New-sheet name (destination === 'newSheet'). */
  sheetName: string;
  /** Top-left anchor for an in-place write (destination === 'existing'). */
  anchor: string;
}

const RANGE_NOTE_STYLE: CSSProperties = {
  fontSize: 12,
  color: 'var(--cs-chrome-muted, #605e5c)',
  marginBottom: 12,
};

const CHECK_STYLE: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 6,
  marginBottom: 12,
  cursor: 'pointer',
};

const RADIO_ROW_STYLE: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 6,
  marginBottom: 6,
  cursor: 'pointer',
};

export function InsertPivotDialog({ api, onClose }: DialogComponentProps) {
  // Read the source grid + its A1 label once on mount.
  const source = useMemo(() => readSource(api), [api]);
  const rangeLabel = useMemo(() => {
    const r = activeRange(api) as unknown as RangeView | null;
    return r?.getA1Notation?.() ?? null;
  }, [api]);

  const hasSource = source.length > 0 && (source[0]?.length ?? 0) > 0;
  const colCount = hasSource ? source[0].length : 0;

  const [state, setState] = useState<DialogState>({
    useHeader: true,
    groupCol: 0,
    valueCol: colCount > 1 ? 1 : 0,
    aggFn: 'sum',
    destination: 'newSheet',
    sheetName: 'Pivot',
    anchor: '',
  });
  const [error, setError] = useState<string | null>(null);

  const update = <K extends keyof DialogState>(key: K, value: DialogState[K]) => {
    setState((prev) => ({ ...prev, [key]: value }));
    setError(null);
  };

  // Column choices, labelled from the header row when "first row is header".
  const columnOptions = useMemo(() => {
    const header = state.useHeader ? source[0] : undefined;
    return Array.from({ length: colCount }, (_, i) => ({
      value: i,
      label: columnLabel(i, header, state.useHeader),
    }));
  }, [source, colCount, state.useHeader]);

  const apply = () => {
    const pivot = buildPivot(source, state.useHeader, state.groupCol, state.valueCol, state.aggFn);
    if (!pivot) {
      setError('The selection has no data rows to summarize.');
      return;
    }

    const wb = api.univer.getActiveWorkbook();
    if (!wb) {
      setError('No active workbook.');
      return;
    }

    // Source range coords (absolute) — the pivot model reads its data from here,
    // and the Fields panel re-applies against the same range on every edit.
    const srcRange = activeRange(api) as unknown as RangeView | null;
    const activeWs = wb.getActiveSheet() as unknown as SheetView | null;
    const srcRow = srcRange?.getRow?.() ?? 0;
    const srcCol = srcRange?.getColumn?.() ?? 0;
    const sourceSheetId = activeWs?.getSheetId?.() ?? '';
    const sourceRange = {
      startRow: srcRow,
      startColumn: srcCol,
      endRow: srcRow + source.length - 1,
      endColumn: srcCol + (source[0]?.length ?? 1) - 1,
    };

    let targetSheetId = sourceSheetId;
    let target = { row: 0, column: 0 };

    try {
      if (state.destination === 'newSheet') {
        const name = state.sheetName.trim() || 'Pivot';
        const rows = Math.max(pivot.grid.length + 2, 10);
        // create(name, rows, columns) → FWorksheet (f-workbook.d.ts L178).
        const sheet = (
          wb as unknown as { create: (n: string, r: number, c: number) => SheetView }
        ).create(name, rows, 4);
        sheet.getRange(`A1:${toA1(pivot.grid.length - 1, 1)}`).setValues(pivot.grid);
        targetSheetId = sheet.getSheetId?.() ?? sourceSheetId;
        target = { row: 0, column: 0 };
      } else {
        const anchor = state.anchor.trim().toUpperCase();
        if (!/^[A-Z]+[0-9]+$/.test(anchor)) {
          setError('Enter a valid top-left cell, e.g. E1.');
          return;
        }
        const sheet = wb.getActiveSheet() as unknown as SheetView | null;
        if (!sheet) {
          setError('No active sheet to write into.');
          return;
        }
        // Anchor + grid extent -> destination A1 range, then setValues.
        const m = /^([A-Z]+)([0-9]+)$/.exec(anchor)!;
        let col = 0;
        for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
        col -= 1;
        const startRow = Number(m[2]) - 1;
        const end = toA1(startRow + pivot.grid.length - 1, col + 1);
        sheet.getRange(`${anchor}:${end}`).setValues(pivot.grid);
        target = { row: startRow, column: col };
      }
    } catch (e) {
      setError(e instanceof Error ? e.message : 'Failed to create the pivot table.');
      return;
    }

    // Persist an editable model (header-based pivots only — the model treats the
    // source's first row as headers) so the Fields panel can reconfigure it and
    // the engine recomputes the output grid on every edit.
    if (state.useHeader && sourceSheetId && targetSheetId) {
      persistPivotModel(api, {
        id: `pivot-${targetSheetId}-${target.row}-${target.column}`,
        sourceSheetId,
        source: sourceRange,
        targetSheetId,
        target,
        rows: [{ column: state.groupCol }],
        cols: [],
        values: [{ column: state.valueCol, agg: state.aggFn }],
        filters: [],
        lastOutputExtent: { rows: pivot.grid.length, cols: pivot.grid[0]?.length ?? 0 },
      });
    }

    onClose();
  };

  return (
    <Dialog
      title="Insert pivot table"
      onClose={onClose}
      width={460}
      data-testid="cs-insert-pivot-dialog"
      footer={
        <>
          <button type="button" style={DIALOG_BTN_SECONDARY_STYLE} onClick={onClose}>
            Cancel
          </button>
          <button
            type="button"
            style={DIALOG_BTN_PRIMARY_STYLE}
            data-testid="cs-insert-pivot-create"
            disabled={!hasSource}
            onClick={apply}
          >
            Create
          </button>
        </>
      }
    >
      {hasSource ? (
        <div style={RANGE_NOTE_STYLE} data-testid="cs-insert-pivot-source">
          Source data <strong>{rangeLabel ?? 'the current selection'}</strong> ({source.length} rows
          × {colCount} columns)
        </div>
      ) : (
        <div style={RANGE_NOTE_STYLE} data-testid="cs-insert-pivot-no-selection">
          Select the source data range first (including its header row), then reopen this dialog.
        </div>
      )}

      <label style={CHECK_STYLE} data-testid="cs-insert-pivot-header-label">
        <input
          type="checkbox"
          data-testid="cs-insert-pivot-header"
          checked={state.useHeader}
          onChange={(e) => update('useHeader', e.target.checked)}
          disabled={!hasSource}
        />
        <span>First row is a header</span>
      </label>

      <label style={DIALOG_FIELD_STYLE}>
        <span style={DIALOG_LABEL_STYLE}>Rows (group by)</span>
        <select
          style={DIALOG_INPUT_STYLE}
          data-testid="cs-insert-pivot-group-col"
          value={state.groupCol}
          disabled={!hasSource}
          onChange={(e) => update('groupCol', Number(e.target.value))}
        >
          {columnOptions.map((opt) => (
            <option key={opt.value} value={opt.value}>
              {opt.label}
            </option>
          ))}
        </select>
      </label>

      <label style={DIALOG_FIELD_STYLE}>
        <span style={DIALOG_LABEL_STYLE}>Values</span>
        <select
          style={DIALOG_INPUT_STYLE}
          data-testid="cs-insert-pivot-value-col"
          value={state.valueCol}
          disabled={!hasSource}
          onChange={(e) => update('valueCol', Number(e.target.value))}
        >
          {columnOptions.map((opt) => (
            <option key={opt.value} value={opt.value}>
              {opt.label}
            </option>
          ))}
        </select>
      </label>

      <label style={DIALOG_FIELD_STYLE}>
        <span style={DIALOG_LABEL_STYLE}>Summarize by</span>
        <select
          style={DIALOG_INPUT_STYLE}
          data-testid="cs-insert-pivot-agg"
          value={state.aggFn}
          disabled={!hasSource}
          onChange={(e) => update('aggFn', e.target.value as AggFn)}
        >
          {AGG_OPTIONS.map((opt) => (
            <option key={opt.value} value={opt.value}>
              {opt.label}
            </option>
          ))}
        </select>
      </label>

      <div style={DIALOG_FIELD_STYLE}>
        <span style={DIALOG_LABEL_STYLE}>Destination</span>
        <label style={RADIO_ROW_STYLE}>
          <input
            type="radio"
            name="cs-insert-pivot-destination"
            data-testid="cs-insert-pivot-dest-new"
            checked={state.destination === 'newSheet'}
            onChange={() => update('destination', 'newSheet')}
          />
          <span>New sheet</span>
        </label>
        <label style={RADIO_ROW_STYLE}>
          <input
            type="radio"
            name="cs-insert-pivot-destination"
            data-testid="cs-insert-pivot-dest-existing"
            checked={state.destination === 'existing'}
            onChange={() => update('destination', 'existing')}
          />
          <span>Existing sheet (choose top-left cell)</span>
        </label>
      </div>

      {state.destination === 'newSheet' ? (
        <label style={DIALOG_FIELD_STYLE}>
          <span style={DIALOG_LABEL_STYLE}>New sheet name</span>
          <input
            style={DIALOG_INPUT_STYLE}
            data-testid="cs-insert-pivot-sheet-name"
            value={state.sheetName}
            placeholder="Pivot"
            onChange={(e) => update('sheetName', e.target.value)}
          />
        </label>
      ) : (
        <label style={DIALOG_FIELD_STYLE}>
          <span style={DIALOG_LABEL_STYLE}>Top-left cell (on the active sheet)</span>
          <input
            style={DIALOG_INPUT_STYLE}
            data-testid="cs-insert-pivot-anchor"
            value={state.anchor}
            placeholder="e.g. E1"
            onChange={(e) => update('anchor', e.target.value)}
          />
        </label>
      )}

      {error && (
        <div
          style={{
            ...RANGE_NOTE_STYLE,
            color: 'var(--cs-chrome-danger, #b3261e)',
            marginBottom: 0,
          }}
          data-testid="cs-insert-pivot-error"
          role="alert"
        >
          {error}
        </div>
      )}
    </Dialog>
  );
}
