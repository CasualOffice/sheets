/**
 * Copyright 2026 Casual Office
 *
 * Licensed under the Apache License, Version 2.0 (the "License").
 */

/**
 * Tables side panel: lists the formatted tables on the active sheet with
 * inline rename, theme swatches, and delete. Ported from the standalone app's
 * TablesPanel — the app-only `useUI`/`useBusy` couplings are replaced by the
 * panel host's `onClose`, and the Univer facade is reached through `api.univer`.
 */
import { useEffect, useMemo, useState, type CSSProperties } from 'react';

import { ensurePluginByName } from '../univer';
import type { PanelComponentProps } from './extensions';
import { Icon } from './Icon';
import { PanelHeader, PanelEmpty } from './panel-shell';

// Univer's six built-in table themes, surfaced as swatches.
const TABLE_THEMES = [
  { id: 'table-default-0', label: 'Indigo', swatch: '#6280F9' },
  { id: 'table-default-1', label: 'Teal', swatch: '#16BDCA' },
  { id: 'table-default-2', label: 'Green', swatch: '#31C48D' },
  { id: 'table-default-3', label: 'Purple', swatch: '#AC94FA' },
  { id: 'table-default-4', label: 'Pink', swatch: '#F17EBB' },
  { id: 'table-default-5', label: 'Red', swatch: '#F98080' },
] as const;
type TableThemeId = (typeof TABLE_THEMES)[number]['id'];

interface TableRange {
  startRow: number;
  startColumn: number;
  endRow: number;
  endColumn: number;
}
interface RawTable {
  id: string;
  name: string;
  subUnitId: string;
  range: TableRange;
}
interface TableInfo extends RawTable {
  styleId: string;
}

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
function toA1(r: TableRange): string {
  const start = `${colLetters(r.startColumn)}${r.startRow + 1}`;
  const end = `${colLetters(r.endColumn)}${r.endRow + 1}`;
  return start === end ? start : `${start}:${end}`;
}

const REFRESH_CMDS = new Set([
  'sheet.command.add-table',
  'sheet.command.delete-table',
  'sheet.command.set-table-config',
  'sheet.mutation.add-table',
  'sheet.mutation.set-table-config',
  'sheet.mutation.delete-table',
  'sheet.operation.set-worksheet-activate',
  'doc.command-replace-snapshot',
]);

/* eslint-disable @typescript-eslint/no-explicit-any */
function readTables(univer: any, sheetId: string | null): TableInfo[] {
  try {
    const wb = univer.getActiveWorkbook?.();
    if (!wb) return [];
    const all: RawTable[] = wb.getTableList?.() ?? [];
    return all
      .filter((t) => (sheetId ? t.subUnitId === sheetId : true))
      .map((t) => ({ ...t, styleId: 'table-default-0' }));
  } catch {
    // Table plugin still registering — the CommandExecuted subscription
    // recomputes once it's ready.
    return [];
  }
}

const header: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 8,
  padding: '10px 12px',
  borderBottom: '1px solid var(--cs-chrome-border, #edeff3)',
};

export function TablesPanel({ api, onClose }: PanelComponentProps) {
  const univer = (api as any).univer;
  const [tables, setTables] = useState<TableInfo[]>([]);
  const [renaming, setRenaming] = useState<{ id: string; draft: string } | null>(null);

  useEffect(() => {
    if (!univer) return;
    let cancelled = false;
    let disp: { dispose?: () => void } | undefined;
    const compute = () => {
      if (cancelled) return;
      const sheetId = univer.getActiveWorkbook?.()?.getActiveSheet?.()?.getSheetId?.() ?? null;
      setTables(readTables(univer, sheetId));
    };
    // The table plugin is lazy-loaded — `getTableList()` resolves SheetTableService
    // and throws if the plugin isn't registered yet, so ensure it before reading.
    void ensurePluginByName('table').then(() => {
      if (cancelled) return;
      compute();
      disp = univer.addEvent(univer.Event.CommandExecuted, (e: { id?: string }) => {
        if (e.id && REFRESH_CMDS.has(e.id)) compute();
      });
    });
    return () => {
      cancelled = true;
      disp?.dispose?.();
    };
  }, [univer]);

  const themesById = useMemo(() => {
    const m = new Map<string, (typeof TABLE_THEMES)[number]>();
    for (const t of TABLE_THEMES) m.set(t.id, t);
    return m;
  }, []);

  const empty = tables.length === 0;

  const onRenameCommit = async (id: string, currentName: string) => {
    if (!renaming || renaming.id !== id) return;
    const next = renaming.draft.trim();
    setRenaming(null);
    if (!next || next === currentName) return;
    const sheet = univer.getActiveWorkbook?.()?.getActiveSheet?.();
    if (!sheet) return;
    await ensurePluginByName('table');
    const ok = await Promise.resolve(sheet.setTableName?.(id, next));
    if (ok === false) console.warn(`[tables] rename rejected: "${next}"`);
  };

  const onPickTheme = async (id: string, themeId: TableThemeId) => {
    const wb = univer.getActiveWorkbook?.();
    if (!wb) return;
    await ensurePluginByName('table');
    univer.executeCommand('sheet.command.set-table-config', {
      unitId: wb.getId(),
      tableId: id,
      theme: themeId,
    });
  };

  const onDelete = async (id: string) => {
    const wb = univer.getActiveWorkbook?.();
    if (!wb) return;
    await ensurePluginByName('table');
    wb.removeTable?.(id);
  };

  return (
    <div
      data-testid="cs-tables-panel"
      style={{ height: '100%', display: 'flex', flexDirection: 'column' }}
    >
      <PanelHeader icon="table" title="Tables" count={tables.length} onClose={onClose} />

      <div style={{ flex: 1, overflow: 'auto', padding: 12 }}>
        {empty ? (
          <PanelEmpty icon="table_rows" title="No tables on this sheet" testId="cs-tables-panel-empty">
            Select your data, then use <strong>Insert → Table</strong> from the menu.
          </PanelEmpty>
        ) : (
          <ul
            style={{
              listStyle: 'none',
              margin: 0,
              padding: 0,
              display: 'flex',
              flexDirection: 'column',
              gap: 8,
            }}
          >
            {tables.map((t) => {
              const isRenaming = renaming?.id === t.id;
              const theme = themesById.get(t.styleId);
              return (
                <li
                  key={t.id}
                  data-testid={`cs-tables-panel-row-${t.id}`}
                  style={{
                    border: '1px solid var(--cs-chrome-border, #edeff3)',
                    borderRadius: 8,
                    padding: 10,
                    display: 'flex',
                    flexDirection: 'column',
                    gap: 6,
                  }}
                >
                  <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                    {isRenaming ? (
                      <input
                        autoFocus
                        value={renaming.draft}
                        onChange={(e) => setRenaming({ id: t.id, draft: e.target.value })}
                        onBlur={() => onRenameCommit(t.id, t.name)}
                        onKeyDown={(e) => {
                          if (e.key === 'Enter') onRenameCommit(t.id, t.name);
                          if (e.key === 'Escape') setRenaming(null);
                        }}
                        style={{ flex: 1, font: 'inherit', padding: '2px 4px' }}
                      />
                    ) : (
                      <button
                        type="button"
                        onClick={() => setRenaming({ id: t.id, draft: t.name })}
                        title="Click to rename"
                        style={{
                          flex: 1,
                          textAlign: 'left',
                          border: 'none',
                          background: 'transparent',
                          cursor: 'pointer',
                          font: 'inherit',
                          fontWeight: 600,
                          color: 'inherit',
                        }}
                      >
                        {t.name}
                      </button>
                    )}
                    <span style={{ opacity: 0.6, fontSize: 12 }}>{toA1(t.range)}</span>
                    <button
                      type="button"
                      aria-label={`Delete table ${t.name}`}
                      title="Delete table"
                      onClick={() => onDelete(t.id)}
                      style={{
                        border: 'none',
                        background: 'transparent',
                        cursor: 'pointer',
                        color: 'inherit',
                      }}
                    >
                      <Icon name="delete" size={16} />
                    </button>
                  </div>
                  <div role="group" aria-label="Table theme" style={{ display: 'flex', gap: 4 }}>
                    {TABLE_THEMES.map((opt) => (
                      <button
                        key={opt.id}
                        type="button"
                        title={opt.label}
                        aria-label={opt.label}
                        aria-pressed={theme?.id === opt.id}
                        onClick={() => onPickTheme(t.id, opt.id)}
                        style={{
                          width: 18,
                          height: 18,
                          borderRadius: 4,
                          cursor: 'pointer',
                          background: opt.swatch,
                          border:
                            theme?.id === opt.id
                              ? '2px solid var(--cs-chrome-active-fg, #0e7490)'
                              : '1px solid rgba(0,0,0,0.15)',
                        }}
                      />
                    ))}
                  </div>
                </li>
              );
            })}
          </ul>
        )}
      </div>
    </div>
  );
}
