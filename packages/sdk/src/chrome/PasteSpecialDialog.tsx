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
 * PasteSpecialDialog — the SDK chrome's built-in Paste Special modal.
 *
 * Copies the DataValidationDialog/FormatCellsDialog structure: reads the active
 * A1 selection off the FUniver facade, gathers a mode via a small form, and
 * applies through real Univer commands / facade calls (no app context).
 *
 * Modes → mechanism (all grounded in the INSTALLED @univerjs/sheets-ui build):
 *   - values   → `sheet.command.paste` with `{ value: 'special-paste-value' }`
 *   - formats  → `sheet.command.paste` with `{ value: 'special-paste-format' }`
 *   - formulas → `sheet.command.paste` with `{ value: 'special-paste-formula' }`
 *
 *     The special-paste hook names live in `PREDEFINED_HOOK_NAME_PASTE`
 *     (sheets-ui/lib/es/index.js: SPECIAL_PASTE_VALUE/FORMAT/FORMULA), and the
 *     dedicated `SheetPasteValueCommand`/`SheetPasteFormatCommand` handlers just
 *     forward `SheetPasteCommand.id` (= 'sheet.command.paste', name verified in
 *     the same build) with those `{ value }` params. We dispatch the base
 *     `sheet.command.paste` with the matching hook value so all three modes go
 *     through one real, installed command — including 'formulas', which has no
 *     dedicated `sheet.command.paste-*` id but IS a valid paste hook.
 *
 *   - transpose → in-place matrix flip via the sheets facade:
 *     `FRange.getValues()` → transpose → `getRange(row, col, cols, rows)` →
 *     `.setValues()`. There is NO transpose paste command anywhere in the
 *     installed sheets/sheets-ui build (grep 'transpose' returns nothing), so
 *     rather than stub it we transpose the CURRENT selection in place using the
 *     real `getValues`/`getRange`/`setValues` facade methods (verified in
 *     @univerjs/sheets/facade f-range.d.ts + f-worksheet.d.ts). See the
 *     `limitations` note: this transposes the selection itself, not clipboard
 *     contents (the facade exposes no clipboard-transpose primitive).
 *
 * Mounted by `<DialogHost>` when `openDialog('paste-special')` is called and no
 * host override is registered.
 */

import { useMemo, useState, type CSSProperties } from 'react';
// Side-effect import: installs the sheets-ui facade extensions and registers
// the clipboard/paste commands on the FUniver facade so
// `executeCommand('sheet.command.paste', …)` resolves at runtime. Mirrors the
// `@univerjs/sheets/facade` side-effect import the SDK's api.ts already does for
// the core sheets mixins.
import '@univerjs/sheets-ui/facade';
import type { DialogComponentProps } from './extensions';
import type { CasualSheetsAPI } from '../sheets/api';
import { Dialog } from './Dialog';
import {
  DIALOG_BTN_PRIMARY_STYLE,
  DIALOG_BTN_SECONDARY_STYLE,
  DIALOG_FIELD_STYLE,
  DIALOG_LABEL_STYLE,
} from './dialog-styles';

/** Paste-special modes this dialog offers. */
type PasteMode = 'values' | 'formats' | 'formulas' | 'transpose';

const MODE_OPTIONS: Array<{ value: PasteMode; label: string; hint: string }> = [
  {
    value: 'values',
    label: 'Values only',
    hint: 'Paste cell values, dropping formatting and formulas.',
  },
  {
    value: 'formats',
    label: 'Formats only',
    hint: 'Paste number formats, fonts, borders and fills — no values.',
  },
  {
    value: 'formulas',
    label: 'Formulas only',
    hint: 'Paste formulas, adjusting relative references.',
  },
  {
    value: 'transpose',
    label: 'Transpose',
    hint: 'Flip the current selection — rows become columns and columns become rows, in place.',
  },
];

/** Paste hook names from sheets-ui `PREDEFINED_HOOK_NAME_PASTE`. */
const PASTE_HOOK: Record<'values' | 'formats' | 'formulas', string> = {
  values: 'special-paste-value',
  formats: 'special-paste-format',
  formulas: 'special-paste-formula',
};

/** The base sheets paste command; the special modes are `{ value: <hook> }`. */
const SHEET_PASTE_COMMAND = 'sheet.command.paste';

/** Minimal shape of the FRange we lean on for the transpose path. */
interface TransposableRange {
  getRow(): number;
  getColumn(): number;
  getWidth(): number;
  getHeight(): number;
  getValues(): unknown[][];
  getA1Notation?(): string;
}

/** The active FRange, or null when there is no selection. */
function activeRange(api: CasualSheetsAPI) {
  return api.univer.getActiveWorkbook()?.getActiveSheet()?.getActiveRange() ?? null;
}

/**
 * Transpose the current selection in place using the sheets facade: read the
 * value matrix, flip it, then write it back into a range whose rows/columns are
 * swapped, anchored at the same top-left cell. Returns false when there is no
 * selection or nothing to transpose.
 */
function transposeSelection(api: CasualSheetsAPI): boolean {
  const sheet = api.univer.getActiveWorkbook()?.getActiveSheet();
  const range = sheet?.getActiveRange() as unknown as TransposableRange | null;
  if (!sheet || !range) return false;

  const rows = range.getHeight();
  const cols = range.getWidth();
  if (rows <= 0 || cols <= 0) return false;

  const values = range.getValues();
  const transposed: unknown[][] = [];
  for (let c = 0; c < cols; c++) {
    const newRow: unknown[] = [];
    for (let r = 0; r < rows; r++) {
      newRow.push(values[r]?.[c] ?? null);
    }
    transposed.push(newRow);
  }

  // New range: same anchor, swapped dimensions (cols rows × rows columns).
  const target = sheet.getRange(range.getRow(), range.getColumn(), cols, rows) as unknown as {
    setValues: (v: unknown[][]) => unknown;
  };
  target.setValues(transposed);
  return true;
}

/**
 * Apply the chosen paste-special mode. Values/formats/formulas dispatch the real
 * `sheet.command.paste` with the matching special-paste hook; transpose flips the
 * selection in place. Returns a promise that resolves to whether it ran.
 */
async function applyPasteSpecial(api: CasualSheetsAPI, mode: PasteMode): Promise<boolean> {
  if (activeRange(api) === null) return false;

  if (mode === 'transpose') {
    return transposeSelection(api);
  }

  return api.executeCommand(SHEET_PASTE_COMMAND, { value: PASTE_HOOK[mode] });
}

const MODE_RADIO_ROW_STYLE: CSSProperties = {
  display: 'flex',
  alignItems: 'flex-start',
  gap: 8,
  padding: '8px 10px',
  border: '1px solid var(--cs-chrome-border, #cdd3db)',
  borderRadius: 6,
  marginBottom: 8,
  cursor: 'pointer',
};

const MODE_RADIO_ROW_ACTIVE_STYLE: CSSProperties = {
  ...MODE_RADIO_ROW_STYLE,
  borderColor: 'var(--cs-chrome-active-fg, #0e7490)',
  background: 'var(--cs-chrome-active-bg, rgba(14, 116, 144, 0.06))',
};

const MODE_LABEL_STYLE: CSSProperties = {
  fontSize: 13,
  fontWeight: 500,
  color: 'var(--cs-chrome-fg, #201f1e)',
};

const MODE_HINT_STYLE: CSSProperties = {
  fontSize: 12,
  color: 'var(--cs-chrome-muted, #605e5c)',
  marginTop: 2,
  lineHeight: 1.35,
};

const RANGE_NOTE_STYLE: CSSProperties = {
  fontSize: 12,
  color: 'var(--cs-chrome-muted, #605e5c)',
  marginBottom: 12,
};

export function PasteSpecialDialog({ api, onClose }: DialogComponentProps) {
  const [mode, setMode] = useState<PasteMode>('values');

  // Read the selection once for the header hint (getA1Notation is verified on
  // the sheets facade FRange — f-range.d.ts).
  const rangeLabel = useMemo(() => {
    const fRange = activeRange(api) as unknown as { getA1Notation?: () => string } | null;
    return fRange?.getA1Notation?.() ?? null;
  }, [api]);

  const hasSelection = activeRange(api) !== null;

  const apply = () => {
    void applyPasteSpecial(api, mode).then((ok) => {
      if (ok) onClose();
    });
  };

  return (
    <Dialog
      title="Paste special"
      onClose={onClose}
      width={440}
      data-testid="cs-paste-special-dialog"
      footer={
        <>
          <button type="button" style={DIALOG_BTN_SECONDARY_STYLE} onClick={onClose}>
            Cancel
          </button>
          <button
            type="button"
            style={DIALOG_BTN_PRIMARY_STYLE}
            data-testid="cs-paste-special-apply"
            disabled={!hasSelection}
            onClick={apply}
          >
            Paste
          </button>
        </>
      }
    >
      {hasSelection ? (
        <div style={RANGE_NOTE_STYLE} data-testid="cs-paste-special-range">
          {mode === 'transpose' ? 'Transposes' : 'Pastes into'}{' '}
          <strong>{rangeLabel ?? 'the current selection'}</strong>
        </div>
      ) : (
        <div style={RANGE_NOTE_STYLE} data-testid="cs-paste-special-no-selection">
          Select the destination cell(s) first, then reopen this dialog.
        </div>
      )}

      <div style={DIALOG_FIELD_STYLE}>
        <span style={DIALOG_LABEL_STYLE}>Paste</span>
        <div role="radiogroup" aria-label="Paste special mode">
          {MODE_OPTIONS.map((opt) => {
            const active = mode === opt.value;
            return (
              <label
                key={opt.value}
                style={active ? MODE_RADIO_ROW_ACTIVE_STYLE : MODE_RADIO_ROW_STYLE}
                data-testid={`cs-paste-special-mode-${opt.value}`}
              >
                <input
                  type="radio"
                  name="cs-paste-special-mode"
                  value={opt.value}
                  checked={active}
                  onChange={() => setMode(opt.value)}
                  style={{ marginTop: 2 }}
                />
                <span>
                  <span style={MODE_LABEL_STYLE}>{opt.label}</span>
                  <span style={MODE_HINT_STYLE}>{opt.hint}</span>
                </span>
              </label>
            );
          })}
        </div>
      </div>
    </Dialog>
  );
}
