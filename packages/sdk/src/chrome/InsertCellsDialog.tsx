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
 * InsertCellsDialog — the SDK chrome's built-in Insert Cells modal.
 *
 * Follows the DataValidationDialog / FormatCellsDialog pattern: reads the active
 * A1 selection off the FUniver facade, offers the four Excel/Sheets shift
 * choices, and applies via real, installed facade methods:
 *
 *   - Shift cells right → `FRange.insertCells(Dimension.COLUMNS)`
 *   - Shift cells down  → `FRange.insertCells(Dimension.ROWS)`
 *   - Entire row        → `FWorksheet.insertRows(startRow, rowCount)`
 *   - Entire column     → `FWorksheet.insertColumns(startColumn, colCount)`
 *
 * The shift-direction ↔ Dimension mapping is grounded in the `insertCells`
 * doc-example in `@univerjs/sheets/lib/types/facade/f-range.d.ts` (line 1484):
 * `insertCells(Dimension.COLUMNS)` pushes existing data to the RIGHT, and
 * `insertCells(Dimension.ROWS)` pushes it DOWN. `Dimension` is verified in
 * `@univerjs/core/lib/types/types/enum/dimension.d.ts` (COLUMNS=0, ROWS=1).
 * `FWorksheet.insertRows` / `insertColumns` are verified in
 * `@univerjs/sheets/lib/types/facade/f-worksheet.d.ts` (lines 355 / 679), and
 * the range extents come from `FRange.getRow/getColumn/getLastRow/getLastColumn`
 * (f-range.d.ts lines 115/141/128/154).
 *
 * Mounted by `<DialogHost>` when `openDialog('insert-cells')` is called and no
 * host override is registered.
 */

import { useMemo, useState, type CSSProperties } from 'react';
import { Dimension } from '@univerjs/core';
import type { DialogComponentProps } from './extensions';
import type { CasualSheetsAPI } from '../sheets/api';
import { Dialog } from './Dialog';
import { DIALOG_BTN_PRIMARY_STYLE, DIALOG_BTN_SECONDARY_STYLE } from './dialog-styles';

/** The four shift choices, matching Excel / Google Sheets "Insert cells". */
type ShiftChoice = 'right' | 'down' | 'row' | 'column';

const SHIFT_OPTIONS: Array<{ value: ShiftChoice; label: string }> = [
  { value: 'right', label: 'Shift cells right' },
  { value: 'down', label: 'Shift cells down' },
  { value: 'row', label: 'Entire row' },
  { value: 'column', label: 'Entire column' },
];

/** The active FRange, or null when there is no selection. */
function activeRange(api: CasualSheetsAPI) {
  return api.univer.getActiveWorkbook()?.getActiveSheet()?.getActiveRange() ?? null;
}

/**
 * Insert cells at the active selection with the chosen shift. Returns false when
 * there is no active range (nothing to insert against).
 */
function applyInsert(api: CasualSheetsAPI, shift: ShiftChoice): boolean {
  const range = activeRange(api);
  if (!range) return false;

  switch (shift) {
    case 'right':
      // COLUMNS dimension pushes existing data to the right (f-range.d.ts:1484).
      range.insertCells(Dimension.COLUMNS);
      return true;
    case 'down':
      // ROWS dimension pushes existing data down.
      range.insertCells(Dimension.ROWS);
      return true;
    case 'row': {
      const sheet = api.univer.getActiveWorkbook()?.getActiveSheet();
      if (!sheet) return false;
      const startRow = range.getRow();
      const rowCount = range.getLastRow() - startRow + 1;
      sheet.insertRows(startRow, rowCount);
      return true;
    }
    case 'column': {
      const sheet = api.univer.getActiveWorkbook()?.getActiveSheet();
      if (!sheet) return false;
      const startColumn = range.getColumn();
      const colCount = range.getLastColumn() - startColumn + 1;
      sheet.insertColumns(startColumn, colCount);
      return true;
    }
  }
}

const RANGE_NOTE_STYLE: CSSProperties = {
  fontSize: 12,
  color: 'var(--cs-chrome-muted, #605e5c)',
  marginBottom: 14,
};

const RADIO_STYLE: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 8,
  marginBottom: 10,
  cursor: 'pointer',
};

export function InsertCellsDialog({ api, onClose }: DialogComponentProps) {
  const [shift, setShift] = useState<ShiftChoice>('down');

  // Read the selection once for the header hint. `getA1Notation` off the live
  // FRange (verified in @univerjs/sheets/facade f-range.d.ts) gives the
  // user-facing A1 label, e.g. "A1:B2".
  const rangeLabel = useMemo(() => {
    const fRange = activeRange(api) as unknown as { getA1Notation?: () => string } | null;
    return fRange?.getA1Notation?.() ?? null;
  }, [api]);

  const hasSelection = activeRange(api) !== null;

  const apply = () => {
    if (applyInsert(api, shift)) onClose();
  };

  return (
    <Dialog
      title="Insert cells"
      onClose={onClose}
      width={380}
      data-testid="cs-insert-cells-dialog"
      footer={
        <>
          <button type="button" style={DIALOG_BTN_SECONDARY_STYLE} onClick={onClose}>
            Cancel
          </button>
          <button
            type="button"
            style={DIALOG_BTN_PRIMARY_STYLE}
            data-testid="cs-insert-cells-apply"
            disabled={!hasSelection}
            onClick={apply}
          >
            Insert
          </button>
        </>
      }
    >
      {hasSelection ? (
        <div style={RANGE_NOTE_STYLE} data-testid="cs-insert-cells-range">
          Insert at <strong>{rangeLabel ?? 'the current selection'}</strong>
        </div>
      ) : (
        <div style={RANGE_NOTE_STYLE} data-testid="cs-insert-cells-no-selection">
          Select one or more cells first, then reopen this dialog.
        </div>
      )}

      <div role="radiogroup" aria-label="Shift direction">
        {SHIFT_OPTIONS.map((opt) => (
          <label
            key={opt.value}
            style={RADIO_STYLE}
            data-testid={`cs-insert-cells-shift-${opt.value}-label`}
          >
            <input
              type="radio"
              name="cs-insert-cells-shift"
              data-testid={`cs-insert-cells-shift-${opt.value}`}
              checked={shift === opt.value}
              onChange={() => setShift(opt.value)}
            />
            <span>{opt.label}</span>
          </label>
        ))}
      </div>
    </Dialog>
  );
}
