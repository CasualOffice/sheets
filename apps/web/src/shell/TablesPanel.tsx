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

import { useEffect, useMemo, useState } from 'react';
import type { FUniver } from '@univerjs/core/facade';
import { useUniverAPI } from '../use-univer';
import { useUI } from '../use-ui';
import { ensurePluginByName } from '@casualoffice/sheets/univer';
import { Icon } from './Icon';
import { TABLE_THEMES, formatAsTable, type TableThemeId } from './tab-actions';
import { useBusy } from '../busy-context';

type TableRange = {
  startRow: number;
  startColumn: number;
  endRow: number;
  endColumn: number;
};

type TableInfo = {
  id: string;
  name: string;
  subUnitId: string;
  range: TableRange;
  styleId: string;
  /** Stable id for the worksheet the table lives on, so we can filter when sheets change. */
};

// Univer's getTableList projection (we only care about a subset).
type RawTable = {
  id: string;
  name: string;
  subUnitId: string;
  range: TableRange;
};

function colLetters(col: number): string {
  let n = col + 1;
  let out = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    out = String.fromCharCode(65 + r) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

function toA1(range: TableRange): string {
  const start = `${colLetters(range.startColumn)}${range.startRow + 1}`;
  const end = `${colLetters(range.endColumn)}${range.endRow + 1}`;
  return start === end ? start : `${start}:${end}`;
}

function readTables(api: FUniver, currentSheetId: string | null): TableInfo[] {
  const wb = api.getActiveWorkbook();
  if (!wb) return [];
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const all: RawTable[] = (wb as any).getTableList?.() ?? [];
  return all
    .filter((t) => (currentSheetId ? t.subUnitId === currentSheetId : true))
    .map((t) => ({
      id: t.id,
      name: t.name,
      subUnitId: t.subUnitId,
      range: t.range,
      styleId: window.__getTableStyleId__?.(t.id) ?? 'table-default-0',
    }));
}

const REFRESH_CMD_PREFIXES = [
  'sheet.command.add-table',
  'sheet.command.delete-table',
  'sheet.command.set-table-config',
  'sheet.mutation.add-table',
  'sheet.mutation.set-table-config',
  'sheet.mutation.delete-table',
  // Sheet switch + workbook swap.
  'sheet.operation.set-worksheet-activate',
  'doc.command-replace-snapshot',
];

const shouldRefresh = (id?: string) => !!id && REFRESH_CMD_PREFIXES.some((p) => id === p);

export function TablesPanel() {
  const api = useUniverAPI();
  const busy = useBusy();
  const ui = useUI();
  const [tables, setTables] = useState<TableInfo[]>([]);
  const [renaming, setRenaming] = useState<{ id: string; draft: string } | null>(null);

  useEffect(() => {
    if (!api) return;
    const compute = () => {
      const sheetId = api.getActiveWorkbook()?.getActiveSheet()?.getSheetId() ?? null;
      setTables(readTables(api, sheetId));
    };
    compute();
    const disp = api.addEvent(api.Event.CommandExecuted, (e) => {
      if (shouldRefresh((e as { id?: string }).id)) compute();
    });
    return () => disp.dispose();
  }, [api]);

  const empty = tables.length === 0;

  const onRenameCommit = async (id: string, currentName: string) => {
    if (!api || !renaming || renaming.id !== id) return;
    const next = renaming.draft.trim();
    setRenaming(null);
    if (!next || next === currentName) return;
    const wb = api.getActiveWorkbook();
    const sheet = wb?.getActiveSheet();
    if (!wb || !sheet) return;
    // Table plugin is lazy-loaded; without the await `setTableName` is
    // `undefined` on the facade until the plugin finishes registering
    // and the optional chain silently swallows the call. The rename
    // appears to do nothing.
    await ensurePluginByName('table');
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const result = (sheet as any).setTableName?.(id, next);
    // `setTableName` returns `false` when validation rejects the name
    // (duplicate / forbidden characters per `customNameCharacterCheck`
    // in sheets-table). Surface that to the console so a "rename did
    // nothing" report has a breadcrumb.
    const ok = await Promise.resolve(result);
    if (ok === false) {
      console.warn(`[tables] rename rejected — invalid or duplicate name: "${next}"`);
    }
  };

  const onPickTheme = async (id: string, themeId: TableThemeId) => {
    if (!api) return;
    const wb = api.getActiveWorkbook();
    if (!wb) return;
    await ensurePluginByName('table');
    api.executeCommand('sheet.command.set-table-config', {
      unitId: wb.getId(),
      tableId: id,
      theme: themeId,
    });
  };

  const onDelete = async (id: string) => {
    if (!api) return;
    const wb = api.getActiveWorkbook();
    if (!wb) return;
    await ensurePluginByName('table');
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (wb as any).removeTable?.(id);
  };

  const themesById = useMemo(() => {
    const map = new Map<string, (typeof TABLE_THEMES)[number]>();
    for (const t of TABLE_THEMES) map.set(t.id, t);
    return map;
  }, []);

  return (
    <aside className="side-panel tables-panel" data-testid="tables-panel">
      <header className="side-panel__header">
        <Icon name="table" size="sm" />
        <h2 className="side-panel__title">Tables</h2>
        {!empty && <span className="side-panel__count">{tables.length}</span>}
        <button
          type="button"
          className="side-panel__close"
          aria-label="Close tables panel"
          onClick={ui.toggleTablesPanel}
        >
          <Icon name="close" size="sm" />
        </button>
      </header>
      <div className="tables-panel__body">
        {empty ? (
          <div className="tables-panel__empty" data-testid="tables-panel-empty">
            <Icon name="table_rows" size="lg" className="tables-panel__empty-icon" />
            <div className="tables-panel__empty-title">No tables on this sheet</div>
            <div className="tables-panel__empty-body">
              Select the cells you want to format, then click below — or use{' '}
              <strong>Insert → Table</strong> from the menu.
            </div>
            <button
              type="button"
              className="btn-primary tables-panel__empty-cta"
              data-testid="tables-panel-empty-cta"
              disabled={!api}
              onClick={() => {
                if (!api) return;
                void busy.runBusy('Creating table…', () => formatAsTable(api, 'table-default-0'));
              }}
            >
              Format selection as Table
            </button>
          </div>
        ) : (
          <ul className="tables-panel__list">
            {tables.map((t) => {
              const isRenaming = renaming?.id === t.id;
              const theme = themesById.get(t.styleId);
              return (
                <li
                  className="tables-panel__row"
                  key={t.id}
                  data-testid={`tables-panel-row-${t.id}`}
                >
                  <div className="tables-panel__name">
                    {isRenaming ? (
                      <input
                        autoFocus
                        className="tables-panel__name-input"
                        value={renaming.draft}
                        onChange={(e) => setRenaming({ id: t.id, draft: e.target.value })}
                        onBlur={() => onRenameCommit(t.id, t.name)}
                        onKeyDown={(e) => {
                          if (e.key === 'Enter') onRenameCommit(t.id, t.name);
                          if (e.key === 'Escape') setRenaming(null);
                        }}
                      />
                    ) : (
                      <button
                        type="button"
                        className="tables-panel__name-btn"
                        onClick={() => setRenaming({ id: t.id, draft: t.name })}
                        title="Click to rename"
                      >
                        {t.name}
                      </button>
                    )}
                  </div>
                  <div className="tables-panel__range">{toA1(t.range)}</div>
                  <div className="tables-panel__themes" role="group" aria-label="Table theme">
                    {TABLE_THEMES.map((themeOpt) => (
                      <button
                        key={themeOpt.id}
                        type="button"
                        className={`tables-panel__swatch${
                          theme?.id === themeOpt.id ? ' tables-panel__swatch--active' : ''
                        }`}
                        style={{ backgroundColor: themeOpt.swatch }}
                        title={themeOpt.label}
                        aria-label={themeOpt.label}
                        aria-pressed={theme?.id === themeOpt.id}
                        onClick={() => onPickTheme(t.id, themeOpt.id)}
                      />
                    ))}
                  </div>
                  <button
                    type="button"
                    className="tables-panel__delete"
                    aria-label={`Delete table ${t.name}`}
                    title="Delete table"
                    onClick={() => onDelete(t.id)}
                  >
                    <Icon name="delete" />
                  </button>
                </li>
              );
            })}
          </ul>
        )}
      </div>
    </aside>
  );
}
