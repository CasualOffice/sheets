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
 * DeleteCellsDialog — the SDK chrome's built-in Delete Cells modal.
 *
 * Excel/Sheets "Delete cells…" dialog: delete the active selection with a
 * shift choice — shift the surviving cells left, shift them up, or delete the
 * selection's entire rows / entire columns.
 *
 * Applied through the FUniver facade (no invented commands):
 *   - shift left     → `FRange.deleteCells(Dimension.COLUMNS)`
 *   - shift up       → `FRange.deleteCells(Dimension.ROWS)`
 *   - entire row     → `FWorksheet.deleteRows(startRow, rowCount)`
 *   - entire column  → `FWorksheet.deleteColumns(startColumn, columnCount)`
 *
 * Grounded in the installed facade types:
 *   - `Dimension` enum (COLUMNS=0, ROWS=1) — @univerjs/core dimension.d.ts.
 *   - `FRange.deleteCells(shiftDimension: Dimension)` — @univerjs/sheets
 *     facade f-range.d.ts L1541; `getRow/getColumn/getLastRow/getLastColumn`
 *     L115/141/128/154; `getA1Notation` L1200.
 *   - `FWorksheet.deleteRows(pos, howMany)` L418 / `deleteColumns(pos, howMany)`
 *     L742 — @univerjs/sheets facade f-worksheet.d.ts.
 *
 * Mounted by `<DialogHost>` when `openDialog('delete-cells')` is called and no
 * host override is registered.
 */

import { useMemo, useState, type CSSProperties } from 'react';
import { Dimension } from '@univerjs/core';
import type { DialogComponentProps } from './extensions';
import type { CasualSheetsAPI } from '../sheets/api';
import { Dialog } from './Dialog';
import { DIALOG_BTN_PRIMARY_STYLE, DIALOG_BTN_SECONDARY_STYLE } from './dialog-styles';

/** How surviving cells move (or what dimension is removed wholesale). */
type ShiftChoice = 'left' | 'up' | 'row' | 'column';

const SHIFT_OPTIONS: Array<{ value: ShiftChoice; label: string }> = [
  { value: 'left', label: 'Shift cells left' },
  { value: 'up', label: 'Shift cells up' },
  { value: 'row', label: 'Entire row' },
  { value: 'column', label: 'Entire column' },
];

/** The active FRange, or null when there is no selection. */
function activeRange(api: CasualSheetsAPI) {
  return api.univer.getActiveWorkbook()?.getActiveSheet()?.getActiveRange() ?? null;
}

/**
 * Apply the delete via the facade. Returns false when there's no active range
 * (or no active sheet for the whole-row / whole-column paths).
 */
function applyDelete(api: CasualSheetsAPI, choice: ShiftChoice): boolean {
  const range = activeRange(api);
  if (!range) return false;

  if (choice === 'left') {
    // Existing data to the right shifts left into the gap.
    range.deleteCells(Dimension.COLUMNS);
    return true;
  }
  if (choice === 'up') {
    // Existing data below shifts up into the gap.
    range.deleteCells(Dimension.ROWS);
    return true;
  }

  const sheet = api.univer.getActiveWorkbook()?.getActiveSheet();
  if (!sheet) return false;

  if (choice === 'row') {
    const start = range.getRow();
    const count = range.getLastRow() - start + 1;
    if (count <= 0) return false;
    sheet.deleteRows(start, count);
    return true;
  }

  // choice === 'column'
  const start = range.getColumn();
  const count = range.getLastColumn() - start + 1;
  if (count <= 0) return false;
  sheet.deleteColumns(start, count);
  return true;
}

const RANGE_NOTE_STYLE: CSSProperties = {
  fontSize: 12,
  color: 'var(--cs-chrome-muted, #605e5c)',
  marginBottom: 12,
};

const OPTION_STYLE: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 8,
  padding: '6px 4px',
  cursor: 'pointer',
};

export function DeleteCellsDialog({ api, onClose }: DialogComponentProps) {
  const [choice, setChoice] = useState<ShiftChoice>('left');

  // Read the selection once for the header hint. `getA1Notation` off the live
  // FRange (verified in @univerjs/sheets/facade f-range.d.ts) gives the
  // user-facing A1 label, e.g. "A1:B2".
  const rangeLabel = useMemo(() => {
    const fRange = activeRange(api) as unknown as { getA1Notation?: () => string } | null;
    return fRange?.getA1Notation?.() ?? null;
  }, [api]);

  const hasSelection = activeRange(api) !== null;

  const apply = () => {
    if (applyDelete(api, choice)) onClose();
  };

  return (
    <Dialog
      title="Delete cells"
      onClose={onClose}
      width={360}
      data-testid="cs-delete-cells-dialog"
      footer={
        <>
          <button type="button" style={DIALOG_BTN_SECONDARY_STYLE} onClick={onClose}>
            Cancel
          </button>
          <button
            type="button"
            style={DIALOG_BTN_PRIMARY_STYLE}
            data-testid="cs-delete-cells-apply"
            disabled={!hasSelection}
            onClick={apply}
          >
            Delete
          </button>
        </>
      }
    >
      {hasSelection ? (
        <div style={RANGE_NOTE_STYLE} data-testid="cs-delete-cells-range">
          Delete <strong>{rangeLabel ?? 'the current selection'}</strong> and…
        </div>
      ) : (
        <div style={RANGE_NOTE_STYLE} data-testid="cs-delete-cells-no-selection">
          Select one or more cells first, then reopen this dialog.
        </div>
      )}

      <div role="radiogroup" aria-label="Delete option">
        {SHIFT_OPTIONS.map((opt) => (
          <label key={opt.value} style={OPTION_STYLE}>
            <input
              type="radio"
              name="cs-delete-cells-shift"
              data-testid={`cs-delete-cells-${opt.value}`}
              value={opt.value}
              checked={choice === opt.value}
              onChange={() => setChoice(opt.value)}
            />
            <span>{opt.label}</span>
          </label>
        ))}
      </div>
    </Dialog>
  );
}
