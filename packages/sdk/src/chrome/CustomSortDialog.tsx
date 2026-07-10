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
 * CustomSortDialog — the SDK chrome's built-in "Sort range" modal.
 *
 * Multi-level sort of the active A1 selection: the user adds one or more sort
 * levels (a column + ascending/descending), plus a "data has header row" toggle
 * that keeps the header in place and labels the column pickers with the header
 * text. Applied through the `@univerjs/sheets-sort` facade — the
 * `FRange.sort(SortColumnSpec | SortColumnSpec[])` extension, where
 * `SortColumnSpec = { column: number; ascending: boolean } | number` and
 * `column` is 0-based **relative to the sorted range's first column**
 * (verified in `@univerjs/sheets-sort/lib/types/facade/f-range.d.ts` L17-49).
 *
 * The facade `sort()` has no first-class "has header" option, so the header row
 * is honoured by sorting a sub-range that starts one row below the selection's
 * top (`FWorksheet.getRange(row, col, numRows, numCols)`, verified in
 * `@univerjs/sheets/lib/types/facade/f-worksheet.d.ts` L258) — the header row is
 * simply excluded from the sorted band, so it stays put.
 *
 * Mounted by `<DialogHost>` when `openDialog('custom-sort')` is called and no
 * host override is registered.
 */

import { useMemo, useState, type CSSProperties } from 'react';
// Side-effect import: installs `FRange.sort()` (and the sort mixin's TS
// augmentation of `@univerjs/sheets/facade`) on the facade prototype. The sort
// plugin registers the command but does NOT import this facade module, so
// without it `range.sort(...)` is undefined at runtime. Mirrors the
// `sheets-data-validation/facade` side-effect import in DataValidationDialog.
import '@univerjs/sheets-sort/facade';
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

/** One sort level: a range-relative column offset + direction. */
interface SortLevel {
  /** 0-based offset from the range's first column. */
  column: number;
  ascending: boolean;
}

/** The `SortColumnSpec` object shape from the sheets-sort facade. */
type SortColumnSpec = { column: number; ascending: boolean };

/** The active FRange, or null when there is no selection. */
function activeRange(api: CasualSheetsAPI) {
  return api.univer.getActiveWorkbook()?.getActiveSheet()?.getActiveRange() ?? null;
}

/** Minimal shape we lean on off the live FRange (grounded in sheets/facade). */
interface RangeGeom {
  getRow(): number;
  getColumn(): number;
  getWidth(): number;
  getHeight(): number;
  getA1Notation?: () => string;
  getValues?: () => Array<Array<unknown>>;
}

/** Convert a 0-based absolute column index to a spreadsheet letter (A, B, …, AA). */
function columnLetter(index: number): string {
  let n = index;
  let label = '';
  do {
    label = String.fromCharCode(65 + (n % 26)) + label;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return label;
}

/** Read geometry + optional header labels off the active range. */
interface RangeInfo {
  /** First (absolute) row of the selection. */
  startRow: number;
  /** First (absolute) column of the selection. */
  startColumn: number;
  /** Number of columns in the selection. */
  width: number;
  /** Number of rows in the selection. */
  height: number;
  /** A1 label for the header note, e.g. "A1:C10". */
  a1: string | null;
  /** Values of the first row (header candidates), if readable. */
  firstRow: Array<unknown> | null;
}

function readRangeInfo(api: CasualSheetsAPI): RangeInfo | null {
  const range = activeRange(api) as unknown as RangeGeom | null;
  if (!range) return null;
  const values = range.getValues?.();
  return {
    startRow: range.getRow(),
    startColumn: range.getColumn(),
    width: range.getWidth(),
    height: range.getHeight(),
    a1: range.getA1Notation?.() ?? null,
    firstRow: values && values.length > 0 ? values[0] : null,
  };
}

/**
 * Apply the multi-level sort via the sheets-sort facade. Returns false when
 * there's no selection or nothing sensible to sort.
 */
function applySort(
  api: CasualSheetsAPI,
  info: RangeInfo,
  levels: SortLevel[],
  hasHeader: boolean,
): boolean {
  const worksheet = api.univer.getActiveWorkbook()?.getActiveSheet();
  if (!worksheet) return false;

  // When the data has a header row, exclude it from the sorted band so it stays
  // pinned at the top. Otherwise sort the whole selection.
  const bandStartRow = hasHeader ? info.startRow + 1 : info.startRow;
  const bandHeight = hasHeader ? info.height - 1 : info.height;
  if (bandHeight <= 0) return false;

  // getRange(row, column, numRows, numColumns) — sheets/facade f-worksheet.d.ts L258.
  const band = (
    worksheet as unknown as {
      getRange: (row: number, col: number, numRows: number, numCols: number) => unknown;
    }
  ).getRange(bandStartRow, info.startColumn, bandHeight, info.width);

  // Clamp each level's column offset to the range and drop invalid ones.
  const specs: SortColumnSpec[] = levels
    .filter((l) => l.column >= 0 && l.column < info.width)
    .map((l) => ({ column: l.column, ascending: l.ascending }));
  if (specs.length === 0) return false;

  // FRange.sort(SortColumnSpec[]) — sheets-sort/facade f-range.d.ts L46. Column
  // indices are relative to the (sub-)range's first column, which is exactly the
  // 0-based offset the pickers produce.
  (band as { sort: (c: SortColumnSpec | SortColumnSpec[]) => unknown }).sort(specs);
  return true;
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

const LEVEL_ROW_STYLE: CSSProperties = {
  display: 'grid',
  gridTemplateColumns: 'auto 1fr 130px 28px',
  alignItems: 'center',
  gap: 8,
  marginBottom: 8,
};

const LEVEL_PREFIX_STYLE: CSSProperties = {
  fontSize: 12,
  color: 'var(--cs-chrome-muted, #605e5c)',
  whiteSpace: 'nowrap',
};

const REMOVE_BTN_STYLE: CSSProperties = {
  width: 28,
  height: 30,
  display: 'inline-flex',
  alignItems: 'center',
  justifyContent: 'center',
  border: '1px solid var(--cs-chrome-border, #cdd3db)',
  borderRadius: 6,
  background: 'var(--cs-chrome-input-bg, #fff)',
  color: 'var(--cs-chrome-muted, #605e5c)',
  cursor: 'pointer',
  padding: 0,
  fontSize: 16,
  lineHeight: 1,
};

const ADD_BTN_STYLE: CSSProperties = {
  ...DIALOG_BTN_SECONDARY_STYLE,
  marginTop: 4,
};

export function CustomSortDialog({ api, onClose }: DialogComponentProps) {
  // Read the selection once when the dialog mounts.
  const info = useMemo(() => readRangeInfo(api), [api]);
  const hasSelection = info !== null;
  const width = info?.width ?? 1;

  const [hasHeader, setHasHeader] = useState(false);
  const [levels, setLevels] = useState<SortLevel[]>([{ column: 0, ascending: true }]);

  // Column option labels: header text when "has header" is on and text exists,
  // otherwise the absolute spreadsheet column letter.
  const columnOptions = useMemo(() => {
    const startCol = info?.startColumn ?? 0;
    return Array.from({ length: width }, (_, offset) => {
      const headerCell = hasHeader ? info?.firstRow?.[offset] : undefined;
      const headerText =
        headerCell !== undefined && headerCell !== null && String(headerCell).trim() !== ''
          ? String(headerCell)
          : null;
      const letter = columnLetter(startCol + offset);
      return {
        value: offset,
        label: headerText ? `${headerText} (${letter})` : `Column ${letter}`,
      };
    });
  }, [width, hasHeader, info]);

  const updateLevel = (index: number, patch: Partial<SortLevel>) =>
    setLevels((prev) => prev.map((l, i) => (i === index ? { ...l, ...patch } : l)));

  const addLevel = () => {
    // Default the new level to the first column not already used, else column 0.
    const used = new Set(levels.map((l) => l.column));
    let next = 0;
    for (let c = 0; c < width; c += 1) {
      if (!used.has(c)) {
        next = c;
        break;
      }
    }
    setLevels((prev) => [...prev, { column: next, ascending: true }]);
  };

  const removeLevel = (index: number) =>
    setLevels((prev) => (prev.length <= 1 ? prev : prev.filter((_, i) => i !== index)));

  const apply = () => {
    if (info && applySort(api, info, levels, hasHeader)) onClose();
  };

  const canApply = hasSelection && levels.length > 0;

  return (
    <Dialog
      title="Sort range"
      onClose={onClose}
      width={480}
      data-testid="cs-custom-sort-dialog"
      footer={
        <>
          <button type="button" style={DIALOG_BTN_SECONDARY_STYLE} onClick={onClose}>
            Cancel
          </button>
          <button
            type="button"
            style={DIALOG_BTN_PRIMARY_STYLE}
            data-testid="cs-custom-sort-apply"
            disabled={!canApply}
            onClick={apply}
          >
            Sort
          </button>
        </>
      }
    >
      {hasSelection ? (
        <div style={RANGE_NOTE_STYLE} data-testid="cs-custom-sort-range">
          Sort <strong>{info?.a1 ?? 'the current selection'}</strong>
        </div>
      ) : (
        <div style={RANGE_NOTE_STYLE} data-testid="cs-custom-sort-no-selection">
          Select the range you want to sort first, then reopen this dialog.
        </div>
      )}

      <label style={CHECK_STYLE} data-testid="cs-custom-sort-has-header-label">
        <input
          type="checkbox"
          data-testid="cs-custom-sort-has-header"
          checked={hasHeader}
          onChange={(e) => setHasHeader(e.target.checked)}
        />
        <span>Data has header row</span>
      </label>

      <div style={DIALOG_FIELD_STYLE}>
        <span style={DIALOG_LABEL_STYLE}>Sort by</span>
        {levels.map((level, index) => (
          <div key={index} style={LEVEL_ROW_STYLE} data-testid={`cs-custom-sort-level-${index}`}>
            <span style={LEVEL_PREFIX_STYLE}>{index === 0 ? 'Sort by' : 'then by'}</span>
            <select
              style={DIALOG_INPUT_STYLE}
              data-testid={`cs-custom-sort-column-${index}`}
              value={level.column}
              onChange={(e) => updateLevel(index, { column: Number(e.target.value) })}
            >
              {columnOptions.map((opt) => (
                <option key={opt.value} value={opt.value}>
                  {opt.label}
                </option>
              ))}
            </select>
            <select
              style={DIALOG_INPUT_STYLE}
              data-testid={`cs-custom-sort-order-${index}`}
              value={level.ascending ? 'asc' : 'desc'}
              onChange={(e) => updateLevel(index, { ascending: e.target.value === 'asc' })}
            >
              <option value="asc">A → Z (ascending)</option>
              <option value="desc">Z → A (descending)</option>
            </select>
            <button
              type="button"
              style={REMOVE_BTN_STYLE}
              data-testid={`cs-custom-sort-remove-${index}`}
              aria-label="Remove sort level"
              disabled={levels.length <= 1}
              onClick={() => removeLevel(index)}
            >
              −
            </button>
          </div>
        ))}

        <button
          type="button"
          style={ADD_BTN_STYLE}
          data-testid="cs-custom-sort-add-level"
          disabled={levels.length >= width}
          onClick={addLevel}
        >
          Add another sort level
        </button>
      </div>
    </Dialog>
  );
}
