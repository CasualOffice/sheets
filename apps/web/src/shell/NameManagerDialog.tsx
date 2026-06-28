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

import { useEffect, useState } from 'react';
import type { FUniver } from '@univerjs/core/facade';
import { Dialog } from './Dialog';
import { Icon } from './Icon';

/**
 * Excel-style Name Manager (Ctrl+F3). Lists every named range in the
 * active workbook with its refers-to string and an inline editor for
 * create / rename / re-point / delete.
 *
 * Reads / writes via Univer's FWorkbook facade
 * (`getDefinedNames`, `insertDefinedName`, `updateDefinedNameBuilder`,
 * `deleteDefinedName`) — the mutations these emit are part of
 * `SYNCED_MUTATIONS`, so changes propagate cross-peer in co-edit and
 * round-trip through xlsx (defined-names section).
 *
 * Names can be workbook-scoped or scoped to a single worksheet (Excel's
 * "Scope" column) via the FDefinedName builder's `setScopeToWorksheet` /
 * `setScopeToWorkbook`. The dialog re-reads the live list on each mutation so
 * the UI stays consistent with the source of truth — no incremental diff
 * bookkeeping.
 */

// Sentinel localSheetId meaning "whole workbook" (SCOPE_WORKBOOK_VALUE_DEFINED_NAME
// in the fork). Hard-coded to avoid importing @univerjs/sheets into the shell.
const WORKBOOK_SCOPE = 'AllDefaultWorkbook';

type Props = {
  api: FUniver;
  onClose: () => void;
};

type SheetRef = { id: string; name: string };

type Entry = {
  name: string;
  ref: string;
  /** localSheetId, or WORKBOOK_SCOPE for a workbook-scoped name. */
  scope: string;
};

export function NameManagerDialog({ api, onClose }: Props) {
  const [entries, setEntries] = useState<Entry[]>([]);
  const [sheets, setSheets] = useState<SheetRef[]>([]);
  const [editing, setEditing] = useState<{
    originalName: string | null;
    name: string;
    ref: string;
    scope: string;
  } | null>(null);
  const [error, setError] = useState<string | null>(null);

  const refresh = () => {
    const wb = api.getActiveWorkbook();
    if (!wb) {
      setEntries([]);
      setSheets([]);
      return;
    }
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const wbAny = wb as any;
    setSheets(
      (wbAny.getSheets?.() ?? []).map((s: { getSheetId(): string; getSheetName(): string }) => ({
        id: s.getSheetId(),
        name: s.getSheetName(),
      })),
    );
    const list = wbAny.getDefinedNames?.() as
      | Array<{
          getName: () => string;
          getFormulaOrRefString: () => string;
          getLocalSheetId?: () => string | undefined;
        }>
      | undefined;
    if (!list) {
      setEntries([]);
      return;
    }
    setEntries(
      list
        .map((d) => ({
          name: d.getName(),
          ref: d.getFormulaOrRefString(),
          scope: d.getLocalSheetId?.() ?? WORKBOOK_SCOPE,
        }))
        .sort((a, b) => a.name.localeCompare(b.name)),
    );
  };

  useEffect(refresh, [api]);

  const scopeLabel = (scope: string): string =>
    scope === WORKBOOK_SCOPE ? 'Workbook' : (sheets.find((s) => s.id === scope)?.name ?? scope);

  const startCreate = () =>
    setEditing({ originalName: null, name: '', ref: '', scope: WORKBOOK_SCOPE });
  const startEdit = (e: Entry) =>
    setEditing({ originalName: e.name, name: e.name, ref: e.ref, scope: e.scope });
  const cancelEdit = () => {
    setEditing(null);
    setError(null);
  };

  const commitEdit = () => {
    if (!editing) return;
    const name = editing.name.trim();
    const ref = editing.ref.trim();
    if (!name) {
      setError('Name is required.');
      return;
    }
    if (!ref) {
      setError('Refers to is required.');
      return;
    }
    // Excel rule: defined names can't start with a digit and can't
    // contain spaces. Loose check — Univer enforces stricter rules
    // server-side and will silently fail on bad names.
    if (/^[0-9]/.test(name) || /\s/.test(name)) {
      setError('Name must start with a letter and contain no spaces.');
      return;
    }
    const wb = api.getActiveWorkbook();
    if (!wb) return;
    const scope = editing.scope;
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const wbAny = wb as any;
      // Apply the chosen scope to a builder (workbook or a specific sheet).
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const withScope = (b: any) => {
        if (scope === WORKBOOK_SCOPE) return b.setScopeToWorkbook();
        const fws = wbAny.getSheetBySheetId(scope);
        return fws ? b.setScopeToWorksheet(fws) : b;
      };
      if (editing.originalName == null) {
        // Create via the builder so the scope is honoured.
        const param = withScope(wbAny.newDefinedNameBuilder().setName(name).setRef(ref)).build();
        wbAny.insertDefinedNameBuilder(param);
      } else {
        // Update (rename / re-point / re-scope) on the same id.
        const existing = wbAny.getDefinedName(editing.originalName);
        if (existing) {
          const param = withScope(existing.toBuilder().setName(name).setRef(ref)).build();
          wbAny.updateDefinedNameBuilder(param);
        }
      }
    } catch (err) {
      console.warn('[name-manager] commit failed', err);
      setError('Could not save — see console for details.');
      return;
    }
    setEditing(null);
    setError(null);
    // Univer's mutation completes synchronously; one microtask is
    // enough for the resource model to settle.
    queueMicrotask(refresh);
  };

  const onDelete = (name: string) => {
    const wb = api.getActiveWorkbook();
    if (!wb) return;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (wb as any).deleteDefinedName?.(name);
    queueMicrotask(refresh);
  };

  return (
    <Dialog
      title="Name Manager"
      onClose={onClose}
      data-testid="name-manager-dialog"
      footer={
        editing ? (
          <>
            <button
              type="button"
              className="btn-secondary"
              data-testid="name-manager-cancel"
              onClick={cancelEdit}
            >
              Cancel
            </button>
            <button
              type="button"
              className="btn-primary"
              data-testid="name-manager-save"
              onClick={commitEdit}
            >
              {editing.originalName == null ? 'Create' : 'Save'}
            </button>
          </>
        ) : (
          <>
            <button
              type="button"
              className="btn-primary"
              data-testid="name-manager-new"
              onClick={startCreate}
            >
              <Icon name="add" size="sm" /> New
            </button>
            <button
              type="button"
              className="btn-secondary"
              data-testid="name-manager-close"
              onClick={onClose}
            >
              Close
            </button>
          </>
        )
      }
    >
      {editing ? (
        <div className="name-manager__form">
          <label className="name-manager__field">
            <span>Name</span>
            <input
              autoFocus
              type="text"
              data-testid="name-manager-name-input"
              value={editing.name}
              onChange={(e) => setEditing({ ...editing, name: e.target.value })}
              onKeyDown={(e) => {
                if (e.key === 'Enter') {
                  e.preventDefault();
                  commitEdit();
                } else if (e.key === 'Escape') {
                  e.preventDefault();
                  cancelEdit();
                }
              }}
              placeholder="MyRange"
            />
          </label>
          <label className="name-manager__field">
            <span>Refers to</span>
            <input
              type="text"
              data-testid="name-manager-ref-input"
              value={editing.ref}
              onChange={(e) => setEditing({ ...editing, ref: e.target.value })}
              onKeyDown={(e) => {
                if (e.key === 'Enter') {
                  e.preventDefault();
                  commitEdit();
                } else if (e.key === 'Escape') {
                  e.preventDefault();
                  cancelEdit();
                }
              }}
              placeholder="Sheet1!$A$1:$B$10"
              spellCheck={false}
            />
          </label>
          <label className="name-manager__field">
            <span>Scope</span>
            <select
              data-testid="name-manager-scope-select"
              value={editing.scope}
              onChange={(e) => setEditing({ ...editing, scope: e.target.value })}
            >
              <option value={WORKBOOK_SCOPE}>Workbook</option>
              {sheets.map((s) => (
                <option key={s.id} value={s.id}>
                  {s.name}
                </option>
              ))}
            </select>
          </label>
          {error && (
            <div className="name-manager__error" data-testid="name-manager-error">
              {error}
            </div>
          )}
        </div>
      ) : entries.length === 0 ? (
        <div className="name-manager__empty">
          No named ranges yet. Click <strong>New</strong> to create one — e.g. <code>SalesQ3</code>{' '}
          referring to <code>Sheet1!$B$2:$B$100</code>.
        </div>
      ) : (
        <table className="name-manager__table">
          <thead>
            <tr>
              <th>Name</th>
              <th>Refers to</th>
              <th>Scope</th>
              <th aria-label="Actions" />
            </tr>
          </thead>
          <tbody>
            {entries.map((e) => (
              <tr key={`${e.scope}:${e.name}`} data-testid="name-manager-row">
                <td className="name-manager__name">{e.name}</td>
                <td className="name-manager__ref">
                  <code>{e.ref}</code>
                </td>
                <td className="name-manager__scope" data-testid="name-manager-scope">
                  {scopeLabel(e.scope)}
                </td>
                <td className="name-manager__actions">
                  <button
                    type="button"
                    className="name-manager__action"
                    data-testid="name-manager-edit"
                    title="Edit"
                    onClick={() => startEdit(e)}
                  >
                    <Icon name="edit" size="sm" />
                  </button>
                  <button
                    type="button"
                    className="name-manager__action name-manager__action--danger"
                    data-testid="name-manager-delete"
                    title="Delete"
                    onClick={() => onDelete(e.name)}
                  >
                    <Icon name="delete" size="sm" />
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      )}
    </Dialog>
  );
}
