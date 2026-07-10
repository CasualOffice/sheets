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
 * NameManagerDialog — the SDK chrome's built-in Name Manager (defined names).
 *
 * Mirrors the Data Validation / Format Cells exemplars: it reaches the workbook
 * off the FUniver facade and drives the workbook's defined-name API to list,
 * add, edit, and delete named ranges. Grounded in
 * `@univerjs/sheets/lib/types/facade/f-workbook.d.ts`:
 *   - `getDefinedNames(): FDefinedName[]`                       (L571)
 *   - `getDefinedName(name): FDefinedName | null`              (L559)
 *   - `insertDefinedName(name, formulaOrRefString): FWorkbook` (L639)
 *   - `deleteDefinedName(name): boolean`                       (L651)
 *   - `updateDefinedNameBuilder(param): void`                  (L626)
 * and `f-defined-name.d.ts`: `FDefinedName.getName()` (L221) /
 * `getFormulaOrRefString()` (L265) / `toBuilder()` (L385); the builder's
 * `setName` (L44) / `setRef` (L74) / `build()` (L175). The "add from selection"
 * ref comes from `FRange.getA1Notation(true)` (sheet-qualified, e.g.
 * `Sheet1!A1:B2` — `f-range.d.ts` L1220), which is exactly the shape
 * `insertDefinedName`'s `formulaOrRefString` accepts.
 *
 * Univer's defined-name API is unscoped by name (no id needed here): to "edit" a
 * name we rebuild it from `getDefinedName(oldName).toBuilder()`, preserving its
 * identity while updating name + ref, then `updateDefinedNameBuilder`.
 *
 * Mounted by `<DialogHost>` when `openDialog('name-manager')` is called and no
 * host override is registered.
 */

import { useCallback, useMemo, useState, type CSSProperties } from 'react';
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

/** Minimal shape of the FDefinedNameBuilder chain we use (grounded in f-defined-name.ts). */
interface FDefinedNameBuilderLike {
  setName(name: string): FDefinedNameBuilderLike;
  setRef(a1Notation: string): FDefinedNameBuilderLike;
  build(): unknown;
}

/** Minimal shape of an FDefinedName we read (grounded in f-defined-name.ts). */
interface FDefinedNameLike {
  getName(): string;
  getFormulaOrRefString(): string;
  toBuilder(): FDefinedNameBuilderLike;
}

/** Minimal shape of the FWorkbook defined-name surface (grounded in f-workbook.d.ts). */
interface FWorkbookLike {
  getDefinedNames(): FDefinedNameLike[];
  getDefinedName(name: string): FDefinedNameLike | null;
  insertDefinedName(name: string, formulaOrRefString: string): unknown;
  deleteDefinedName(name: string): boolean;
  updateDefinedNameBuilder(param: unknown): void;
  getActiveSheet(): {
    getActiveRange(): { getA1Notation(withSheet?: boolean): string } | null;
  } | null;
}

interface NameRow {
  name: string;
  refersTo: string;
}

/** Names must start with a letter/underscore and avoid A1-like tokens; keep it
 *  loose but block the common mistakes (spaces, leading digit, empty). */
const NAME_RE = /^[A-Za-z_][A-Za-z0-9_.]*$/;

function workbook(api: CasualSheetsAPI): FWorkbookLike | null {
  return (api.univer.getActiveWorkbook() as unknown as FWorkbookLike | null) ?? null;
}

function readRows(api: CasualSheetsAPI): NameRow[] {
  const wb = workbook(api);
  if (!wb) return [];
  return wb.getDefinedNames().map((dn) => ({
    name: dn.getName(),
    refersTo: dn.getFormulaOrRefString(),
  }));
}

/** Sheet-qualified A1 of the current selection, e.g. `Sheet1!A1:B2`, or ''. */
function selectionRef(api: CasualSheetsAPI): string {
  const range = workbook(api)?.getActiveSheet()?.getActiveRange();
  return range?.getA1Notation(true) ?? '';
}

const ROW_STYLE: CSSProperties = {
  display: 'grid',
  gridTemplateColumns: '1fr 1.3fr auto',
  alignItems: 'center',
  gap: 8,
  padding: '6px 8px',
  borderBottom: '1px solid var(--cs-chrome-border, #edeff3)',
};

const LIST_STYLE: CSSProperties = {
  border: '1px solid var(--cs-chrome-border, #cdd3db)',
  borderRadius: 6,
  maxHeight: 200,
  overflow: 'auto',
  marginBottom: 16,
};

const EMPTY_STYLE: CSSProperties = {
  padding: '16px 8px',
  fontSize: 13,
  color: 'var(--cs-chrome-muted, #605e5c)',
  textAlign: 'center',
};

const REF_CELL_STYLE: CSSProperties = {
  fontSize: 12,
  color: 'var(--cs-chrome-muted, #605e5c)',
  overflow: 'hidden',
  textOverflow: 'ellipsis',
  whiteSpace: 'nowrap',
};

const ROW_ACTIONS_STYLE: CSSProperties = {
  display: 'flex',
  gap: 4,
};

const SMALL_BTN_STYLE: CSSProperties = {
  ...DIALOG_BTN_SECONDARY_STYLE,
  height: 24,
  padding: '0 8px',
  fontSize: 12,
};

const SECTION_TITLE_STYLE: CSSProperties = {
  fontSize: 13,
  fontWeight: 600,
  color: 'var(--cs-chrome-fg, #201f1e)',
  margin: '0 0 8px',
};

const ERROR_STYLE: CSSProperties = {
  fontSize: 12,
  color: 'var(--cs-chrome-danger, #b91c1c)',
  marginBottom: 8,
};

const TWO_COL_STYLE: CSSProperties = {
  display: 'grid',
  gridTemplateColumns: '1fr 1.3fr',
  gap: 8,
};

export function NameManagerDialog({ api, onClose }: DialogComponentProps) {
  const [rows, setRows] = useState<NameRow[]>(() => readRows(api));
  // When editing, this holds the ORIGINAL name of the row being edited (used to
  // locate + update it). null means we're adding a new name.
  const [editingOriginal, setEditingOriginal] = useState<string | null>(null);
  const [nameInput, setNameInput] = useState('');
  const [refInput, setRefInput] = useState('');
  const [error, setError] = useState<string | null>(null);

  const initialSelRef = useMemo(() => selectionRef(api), [api]);

  const refresh = useCallback(() => setRows(readRows(api)), [api]);

  const resetForm = useCallback(() => {
    setEditingOriginal(null);
    setNameInput('');
    setRefInput('');
    setError(null);
  }, []);

  const startEdit = (row: NameRow) => {
    setEditingOriginal(row.name);
    setNameInput(row.name);
    setRefInput(row.refersTo);
    setError(null);
  };

  const useSelection = () => {
    const ref = selectionRef(api);
    if (ref) setRefInput(ref);
  };

  const remove = (row: NameRow) => {
    const wb = workbook(api);
    if (!wb) return;
    wb.deleteDefinedName(row.name);
    if (editingOriginal === row.name) resetForm();
    refresh();
  };

  const submit = () => {
    const wb = workbook(api);
    if (!wb) {
      setError('No active workbook.');
      return;
    }
    const name = nameInput.trim();
    const ref = refInput.trim();
    if (!name) {
      setError('Enter a name.');
      return;
    }
    if (!NAME_RE.test(name)) {
      setError('Names must start with a letter or underscore and contain no spaces.');
      return;
    }
    if (!ref) {
      setError('Enter a range or formula (e.g. Sheet1!A1:B2).');
      return;
    }

    // Duplicate-name guard: allow keeping the same name while editing, but block
    // colliding with a different existing name.
    const collides = rows.some(
      (r) => r.name.toLowerCase() === name.toLowerCase() && r.name !== editingOriginal,
    );
    if (collides) {
      setError(`A defined name "${name}" already exists.`);
      return;
    }

    if (editingOriginal === null) {
      // Add.
      wb.insertDefinedName(name, ref);
    } else {
      // Edit — rebuild from the existing name to preserve identity/scope, then
      // update. If the row vanished (e.g. deleted elsewhere) fall back to insert.
      const existing = wb.getDefinedName(editingOriginal);
      if (existing) {
        const param = existing.toBuilder().setName(name).setRef(ref).build();
        wb.updateDefinedNameBuilder(param);
      } else {
        wb.insertDefinedName(name, ref);
      }
    }

    refresh();
    resetForm();
  };

  const isEditing = editingOriginal !== null;

  return (
    <Dialog
      title="Name manager"
      onClose={onClose}
      width={480}
      data-testid="cs-name-manager-dialog"
      footer={
        <>
          <span style={{ flex: 1 }} />
          <button type="button" style={DIALOG_BTN_SECONDARY_STYLE} onClick={onClose}>
            Close
          </button>
        </>
      }
    >
      <h3 style={SECTION_TITLE_STYLE}>Defined names</h3>
      <div style={LIST_STYLE} data-testid="cs-name-manager-list">
        {rows.length === 0 ? (
          <div style={EMPTY_STYLE} data-testid="cs-name-manager-empty">
            No defined names yet. Add one below.
          </div>
        ) : (
          rows.map((row) => (
            <div key={row.name} style={ROW_STYLE} data-testid="cs-name-manager-row">
              <span style={{ fontWeight: 500 }}>{row.name}</span>
              <span style={REF_CELL_STYLE} title={row.refersTo}>
                {row.refersTo}
              </span>
              <span style={ROW_ACTIONS_STYLE}>
                <button
                  type="button"
                  style={SMALL_BTN_STYLE}
                  data-testid="cs-name-manager-edit"
                  onClick={() => startEdit(row)}
                >
                  Edit
                </button>
                <button
                  type="button"
                  style={SMALL_BTN_STYLE}
                  data-testid="cs-name-manager-delete"
                  onClick={() => remove(row)}
                >
                  Delete
                </button>
              </span>
            </div>
          ))
        )}
      </div>

      <h3 style={SECTION_TITLE_STYLE}>{isEditing ? 'Edit name' : 'Add a name'}</h3>

      {error && (
        <div style={ERROR_STYLE} data-testid="cs-name-manager-error">
          {error}
        </div>
      )}

      <div style={TWO_COL_STYLE}>
        <label style={DIALOG_FIELD_STYLE}>
          <span style={DIALOG_LABEL_STYLE}>Name</span>
          <input
            style={DIALOG_INPUT_STYLE}
            data-testid="cs-name-manager-name"
            value={nameInput}
            placeholder="MyRange"
            onChange={(e) => setNameInput(e.target.value)}
          />
        </label>
        <label style={DIALOG_FIELD_STYLE}>
          <span style={DIALOG_LABEL_STYLE}>Refers to</span>
          <input
            style={DIALOG_INPUT_STYLE}
            data-testid="cs-name-manager-ref"
            value={refInput}
            placeholder={initialSelRef || 'Sheet1!A1:B2'}
            onChange={(e) => setRefInput(e.target.value)}
          />
        </label>
      </div>

      <div style={{ display: 'flex', gap: 8, marginTop: 4 }}>
        <button
          type="button"
          style={DIALOG_BTN_SECONDARY_STYLE}
          data-testid="cs-name-manager-use-selection"
          onClick={useSelection}
        >
          Use selection
        </button>
        <span style={{ flex: 1 }} />
        {isEditing && (
          <button
            type="button"
            style={DIALOG_BTN_SECONDARY_STYLE}
            data-testid="cs-name-manager-cancel-edit"
            onClick={resetForm}
          >
            Cancel
          </button>
        )}
        <button
          type="button"
          style={DIALOG_BTN_PRIMARY_STYLE}
          data-testid="cs-name-manager-submit"
          onClick={submit}
        >
          {isEditing ? 'Update' : 'Add'}
        </button>
      </div>
    </Dialog>
  );
}
