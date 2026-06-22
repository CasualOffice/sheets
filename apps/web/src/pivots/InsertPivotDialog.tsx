import { useEffect, useMemo, useRef, useState } from 'react';
import { Dialog } from '../shell/Dialog';
import type { FUniver } from '@univerjs/core/facade';
import { PIVOT_AGG_LABELS, type PivotAggregation, type PivotFilter } from './types';

type Props = {
  api: FUniver;
  /** A1 reference seeded from the active selection. */
  defaultSourceA1: string;
  onCancel: () => void;
  onConfirm: (args: {
    source: { startRow: number; endRow: number; startColumn: number; endColumn: number };
    target: { row: number; column: number };
    /** Outermost row field first; an empty array means Grand-Total-only. */
    rowFieldColumns: number[];
    /** Column fields → cross-tab / matrix layout. Empty means no column
     *  field (the classic single-column-per-value layout). P2 ships a
     *  single column field; the array shape leaves room for nesting. */
    colFieldColumns: number[];
    valueFieldColumn: number;
    aggregation: PivotAggregation;
    filters: PivotFilter[];
  }) => void;
};

/**
 * Excel-style "Create PivotTable" + "PivotTable Fields" dialog,
 * compressed to a single screen. The user picks:
 *
 *   - Source range (pre-filled from the active selection).
 *   - Target cell (A1) — where the pivot's top-left will land.
 *   - Row field (+ optional sub-row field for compact multi-row).
 *   - Column field (optional) → cross-tab / matrix layout.
 *   - Value field + aggregation (Sum / Count / Average / Min / Max).
 *   - Filter field (optional).
 *
 * Headers are derived from the source's first row as soon as the user
 * tabs away from the source input, so the row/value pickers populate
 * with real column names — same instant feedback Excel gives.
 *
 * The full drag-and-drop field list (Excel's PivotTable Fields pane) is
 * still deferred — this single-screen dialog covers the common cases.
 */
export function InsertPivotDialog({ api, defaultSourceA1, onCancel, onConfirm }: Props) {
  const [sourceA1, setSourceA1] = useState(defaultSourceA1);
  const [targetA1, setTargetA1] = useState(suggestTarget(defaultSourceA1));
  const [rowField, setRowField] = useState<number>(0);
  // P1.5 — optional second row field for compact-layout multi-row
  // grouping. -1 means "no sub-row". A full drag-and-drop field list
  // (Excel's PivotTable Fields pane) is still deferred.
  const [subRowField, setSubRowField] = useState<number>(-1);
  // P2 — optional column field → cross-tab / matrix layout. -1 means
  // "no column field" (classic row-only layout). When set, the value
  // field fans out across one column per distinct value of this field.
  const [colField, setColField] = useState<number>(-1);
  const [valueField, setValueField] = useState<number>(1);
  const [aggregation, setAggregation] = useState<PivotAggregation>('sum');
  // P1 — filter field. -1 means "no filter".
  const [filterField, setFilterField] = useState<number>(-1);
  // Set of currently-allowed values for the filter field. Empty = none
  // selected = nothing passes; defaults to "all values" when a filter
  // field is first chosen.
  const [filterAllowed, setFilterAllowed] = useState<Set<string>>(new Set());
  const [error, setError] = useState<string | null>(null);
  const inputRef = useRef<HTMLInputElement>(null);

  const activeSheet = useMemo(() => {
    const wb = api.getActiveWorkbook();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return wb?.getActiveSheet() as any;
  }, [api]);

  // Headers are pulled live from the resolved source range so the
  // pickers always reflect what the user will actually pivot.
  const headers = useMemo<string[]>(() => {
    const range = parseRange(activeSheet, sourceA1);
    if (!range || !activeSheet) return [];
    const out: string[] = [];
    for (let c = range.startColumn; c <= range.endColumn; c++) {
      const v = activeSheet.getRange(range.startRow, c).getValue();
      out.push(v == null || v === '' ? `Column ${c - range.startColumn + 1}` : String(v));
    }
    return out;
  }, [activeSheet, sourceA1]);

  // Distinct values for the chosen filter field. Recomputed when the
  // filter column or source range changes. Kept as a sorted string
  // list so the checklist order is deterministic.
  const filterValues = useMemo<string[]>(() => {
    if (filterField < 0) return [];
    const range = parseRange(activeSheet, sourceA1);
    if (!range || !activeSheet) return [];
    const col = range.startColumn + filterField;
    const seen = new Set<string>();
    for (let r = range.startRow + 1; r <= range.endRow; r += 1) {
      const v = activeSheet.getRange(r, col).getValue();
      const key = v == null ? '' : String(v);
      seen.add(key);
    }
    return [...seen].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
  }, [activeSheet, sourceA1, filterField]);

  // Default to "all values allowed" whenever the filter field changes
  // or the source range produces a new set of distinct values.
  useEffect(() => {
    if (filterField < 0) {
      setFilterAllowed(new Set());
      return;
    }
    setFilterAllowed(new Set(filterValues));
  }, [filterField, filterValues]);

  const confirm = () => {
    const source = parseRange(activeSheet, sourceA1.trim());
    if (!source) {
      setError("Source range doesn't look right — try A1:C10.");
      inputRef.current?.focus();
      return;
    }
    const rows = source.endRow - source.startRow + 1;
    const cols = source.endColumn - source.startColumn + 1;
    if (rows < 2 || cols < 1) {
      setError('Source needs a header row plus at least one data row.');
      return;
    }
    const target = parseCell(activeSheet, targetA1.trim());
    if (!target) {
      setError("Target cell doesn't look right — try E1.");
      return;
    }
    if (rowField < 0 || rowField >= cols) {
      setError('Pick a row field from the source columns.');
      return;
    }
    if (valueField < 0 || valueField >= cols) {
      setError('Pick a value field from the source columns.');
      return;
    }
    // Only ship a filter when the user picked a column AND restricted
    // values — leaving every value checked is equivalent to "no
    // filter" and avoids storing redundant state on the model.
    const filters: PivotFilter[] = [];
    if (filterField >= 0 && filterAllowed.size > 0 && filterAllowed.size < filterValues.length) {
      filters.push({
        column: filterField,
        allowedValues: [...filterAllowed],
      });
    }
    // Compose the row-field list. Outer first; the sub-row only joins
    // when the user picked a different column (picking the same column
    // would just duplicate the indent levels with identical labels).
    const rowFieldColumns: number[] = [rowField];
    if (subRowField >= 0 && subRowField !== rowField) {
      rowFieldColumns.push(subRowField);
    }
    // Compose the column-field list. A column field must differ from the
    // row fields — using the same column on both axes produces a
    // degenerate matrix (every off-diagonal cell empty), so we drop it.
    const colFieldColumns: number[] = [];
    if (colField >= 0 && !rowFieldColumns.includes(colField)) {
      colFieldColumns.push(colField);
    }
    onConfirm({
      source,
      target,
      rowFieldColumns,
      colFieldColumns,
      valueFieldColumn: valueField,
      aggregation,
      filters,
    });
  };

  return (
    <Dialog
      title="Insert PivotTable"
      onClose={onCancel}
      data-testid="insert-pivot-dialog"
      footer={
        <>
          <button
            type="button"
            className="btn-secondary"
            data-testid="insert-pivot-cancel"
            onClick={onCancel}
          >
            Cancel
          </button>
          <button
            type="button"
            className="btn-primary"
            data-testid="insert-pivot-confirm"
            onClick={confirm}
          >
            Insert
          </button>
        </>
      }
    >
      <div className="insert-pivot">
        <fieldset className="insert-pivot__group">
          <legend className="insert-pivot__legend">Source data</legend>
          <input
            ref={inputRef}
            type="text"
            className="insert-pivot__range"
            data-testid="insert-pivot-range"
            value={sourceA1}
            spellCheck={false}
            placeholder="A1:C10"
            onChange={(e) => {
              setSourceA1(e.target.value);
              if (error) setError(null);
            }}
          />
          <p className="insert-pivot__hint">
            First row is treated as headers — the field pickers below populate from the column
            names.
          </p>
        </fieldset>

        <fieldset className="insert-pivot__group">
          <legend className="insert-pivot__legend">Target cell</legend>
          <input
            type="text"
            className="insert-pivot__range"
            data-testid="insert-pivot-target"
            value={targetA1}
            spellCheck={false}
            placeholder="E1"
            onChange={(e) => {
              setTargetA1(e.target.value);
              if (error) setError(null);
            }}
          />
          <p className="insert-pivot__hint">
            Top-left cell of the pivot output. The pivot will overwrite any existing values starting
            here — pick an empty area.
          </p>
        </fieldset>

        <div className="insert-pivot__fields">
          <label className="insert-pivot__field">
            <span className="insert-pivot__field-label">Row field</span>
            <select
              className="insert-pivot__select"
              data-testid="insert-pivot-row-field"
              value={rowField}
              onChange={(e) => setRowField(Number(e.target.value))}
              disabled={headers.length === 0}
            >
              {headers.map((h, i) => (
                <option key={i} value={i}>
                  {h}
                </option>
              ))}
            </select>
          </label>
          <label className="insert-pivot__field">
            <span className="insert-pivot__field-label">Sub-row field</span>
            <select
              className="insert-pivot__select"
              data-testid="insert-pivot-sub-row-field"
              value={subRowField}
              onChange={(e) => setSubRowField(Number(e.target.value))}
              disabled={headers.length === 0}
            >
              <option value={-1}>— None —</option>
              {headers.map((h, i) => (
                <option key={i} value={i}>
                  {h}
                </option>
              ))}
            </select>
          </label>
          <label className="insert-pivot__field">
            <span className="insert-pivot__field-label">Column field</span>
            <select
              className="insert-pivot__select"
              data-testid="insert-pivot-col-field"
              value={colField}
              onChange={(e) => setColField(Number(e.target.value))}
              disabled={headers.length === 0}
            >
              <option value={-1}>— None —</option>
              {headers.map((h, i) => (
                <option key={i} value={i}>
                  {h}
                </option>
              ))}
            </select>
          </label>
          <label className="insert-pivot__field">
            <span className="insert-pivot__field-label">Value field</span>
            <select
              className="insert-pivot__select"
              data-testid="insert-pivot-value-field"
              value={valueField}
              onChange={(e) => setValueField(Number(e.target.value))}
              disabled={headers.length === 0}
            >
              {headers.map((h, i) => (
                <option key={i} value={i}>
                  {h}
                </option>
              ))}
            </select>
          </label>
          <label className="insert-pivot__field">
            <span className="insert-pivot__field-label">Aggregation</span>
            <select
              className="insert-pivot__select"
              data-testid="insert-pivot-aggregation"
              value={aggregation}
              onChange={(e) => setAggregation(e.target.value as PivotAggregation)}
            >
              {(Object.keys(PIVOT_AGG_LABELS) as PivotAggregation[]).map((a) => (
                <option key={a} value={a}>
                  {PIVOT_AGG_LABELS[a]}
                </option>
              ))}
            </select>
          </label>
        </div>

        <fieldset className="insert-pivot__group">
          <legend className="insert-pivot__legend">Filter (optional)</legend>
          <label className="insert-pivot__field">
            <span className="insert-pivot__field-label">Filter field</span>
            <select
              className="insert-pivot__select"
              data-testid="insert-pivot-filter-field"
              value={filterField}
              onChange={(e) => setFilterField(Number(e.target.value))}
              disabled={headers.length === 0}
            >
              <option value={-1}>— None —</option>
              {headers.map((h, i) => (
                <option key={i} value={i}>
                  {h}
                </option>
              ))}
            </select>
          </label>
          {filterField >= 0 && filterValues.length > 0 && (
            <div className="insert-pivot__filter-values" data-testid="insert-pivot-filter-values">
              {filterValues.map((v) => (
                <label key={v} className="insert-pivot__filter-value">
                  <input
                    type="checkbox"
                    data-testid={`insert-pivot-filter-${v || 'blank'}`}
                    checked={filterAllowed.has(v)}
                    onChange={(e) => {
                      const next = new Set(filterAllowed);
                      if (e.target.checked) next.add(v);
                      else next.delete(v);
                      setFilterAllowed(next);
                    }}
                  />
                  <span>{v === '' ? '(blank)' : v}</span>
                </label>
              ))}
            </div>
          )}
        </fieldset>

        {error && (
          <div className="insert-pivot__error" data-testid="insert-pivot-error">
            {error}
          </div>
        )}
      </div>
    </Dialog>
  );
}

/** Default target = the cell two columns to the right of the source. */
function suggestTarget(sourceA1: string): string {
  const m = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/i.exec(sourceA1.trim());
  if (!m) return 'E1';
  const endCol = (m[3] ?? m[1]).toUpperCase();
  const startRow = Number(m[2]);
  return `${shiftCol(endCol, 2)}${startRow}`;
}

function shiftCol(letters: string, by: number): string {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  n += by;
  let out = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out || 'A';
}

function parseRange(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  ws: any,
  a1: string,
): {
  startRow: number;
  endRow: number;
  startColumn: number;
  endColumn: number;
} | null {
  if (!ws || !a1) return null;
  try {
    const r = ws.getRange(a1);
    const raw = r?.getRange?.();
    if (!raw) return null;
    return {
      startRow: raw.startRow,
      endRow: raw.endRow,
      startColumn: raw.startColumn,
      endColumn: raw.endColumn,
    };
  } catch {
    return null;
  }
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function parseCell(ws: any, a1: string): { row: number; column: number } | null {
  const r = parseRange(ws, a1);
  if (!r) return null;
  return { row: r.startRow, column: r.startColumn };
}
