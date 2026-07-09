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
 * MenuBar — the built-in dropdown menu row for `<CasualSheets chrome>`.
 *
 * Office / Google-Sheets-style horizontal menu strip (no logo, no title — the
 * host frames the editor with its own bar): File · Edit · View · Insert ·
 * Format · Data · Help. Each top-level button opens a dropdown of items that
 * either dispatch a verified Univer command / facade call through the
 * `CasualSheetsAPI` (no app context), or — for actions that need a dialog the
 * SDK doesn't ship — route through the optional `onDialogRequest` host hook.
 *
 * This mirrors `apps/web/src/shell/MenuBar.tsx`, but reimplements every action
 * INLINE against `api` / `api.univer` (the FUniver facade) + `api.executeCommand`.
 * The command ids + facade logic are copied verbatim from
 * `apps/web/src/shell/home-tab-actions.ts` + `tab-actions.ts` so behaviour
 * matches the real app exactly.
 *
 * Feature flags: pass `features` to hide a control or whole menu group when its
 * feature is disabled. Defaults to all-enabled. A control whose feature is
 * `false` does not render. An entire top-level menu that ends up with no
 * runnable items is dropped.
 *
 * Dialog routing: items marked `dialog: '<kind>'` (Format Cells, Insert Chart,
 * PivotTable, Find & Replace, Insert/Delete cells, Sparkline, Name Manager,
 * Goal Seek, Data Validation, Conditional Formatting, …) call
 * `onDialogRequest(kind, context)` so the host renders its own UI. When
 * `onDialogRequest` is not provided, those items are omitted (the SDK never
 * fakes a dialog).
 *
 * Self-contained: only one menu is open at a time; Escape and an outside
 * pointerdown close it.
 */

import { useEffect, useRef, useState, type CSSProperties, type ReactNode } from 'react';
import { BorderStyleTypes, BorderType } from '@univerjs/core';
import { openExternal } from './openExternal';
import type { FUniver } from '@univerjs/core/facade';
import type { CasualSheetsAPI } from '../sheets/api';
import { ensurePluginByName } from '../univer';
import { Icon } from './Icon';
import { ensureChromeFonts } from './fonts';
import { useDialogs } from './dialog-context';
import type { ChromeExtensions } from './extensions';
import {
  computeVisibleMenus,
  type MenuDef,
  type MenuId,
  type MenuItemDef,
  type RunFn,
} from './menu-model';

export type { MenuDialogKind } from './menu-model';

/**
 * The menu data model + pure gating engine live in `./menu-model` (univer-free,
 * so unit-testable). `RunFn` / `MenuItemDef` / `MenuDef` / `MenuId` /
 * `MenuDialogKind` are re-imported above; the `run` handlers that populate
 * `MENUS` are defined inline below against the `@univerjs/core` facade.
 */

/**
 * Pretty-print a `Ctrl+Shift+X` shortcut for the current platform. The SDK
 * doesn't import the app's `formatShortcut`, so this is a compact inline
 * equivalent: on macOS, swap Ctrl→⌘, Alt→⌥, Shift→⇧ and drop the `+`.
 */
function fmtShortcut(shortcut: string): string {
  const isMac =
    typeof navigator !== 'undefined' && /Mac|iPhone|iPad/i.test(navigator.platform || '');
  if (!isMac) return shortcut;
  return shortcut
    .split('+')
    .map((part) => {
      switch (part) {
        case 'Ctrl':
          return '⌘';
        case 'Alt':
          return '⌥';
        case 'Shift':
          return '⇧';
        default:
          return part;
      }
    })
    .join('');
}

/* ───────────────────────── inline action helpers ─────────────────────────
 * Every helper below is a faithful port of the corresponding function in
 * apps/web/src/shell/home-tab-actions.ts + tab-actions.ts, reimplemented over
 * the SDK's `api` (CasualSheetsAPI) / `api.univer` (FUniver) so the chrome
 * never imports app context. Command ids are identical to the app's.
 */

function fu(api: CasualSheetsAPI): FUniver {
  return api.univer;
}

function activeRange(api: CasualSheetsAPI) {
  return fu(api).getActiveWorkbook()?.getActiveSheet()?.getActiveRange() ?? null;
}

function activeSheet(api: CasualSheetsAPI) {
  return fu(api).getActiveWorkbook()?.getActiveSheet() ?? null;
}

/** Selection as A1 (e.g. `B2:D10`), or null. Used to seed dialog context. */
function selectionA1(api: CasualSheetsAPI): string | null {
  const range = activeRange(api);
  if (!range) return null;
  try {
    return range.getA1Notation();
  } catch {
    return null;
  }
}

/* ── Edit ─────────────────────────────────────────────────────────────── */

const undo: RunFn = (api) => void api.executeCommand('univer.command.undo');
const redo: RunFn = (api) => void api.executeCommand('univer.command.redo');
const cut: RunFn = (api) => void api.executeCommand('univer.command.cut');
const copy: RunFn = (api) => void api.executeCommand('univer.command.copy');
const paste: RunFn = (api) => void api.executeCommand('univer.command.paste');
const pasteFormattingOnly: RunFn = (api) => void api.executeCommand('sheet.command.paste-format');

/* ── Insert (facade-driven; mirror tab-actions.ts) ─────────────────────── */

const insertRowAbove: RunFn = (api) => {
  const range = activeRange(api);
  const sheet = activeSheet(api);
  if (!range || !sheet) return;
  sheet.insertRowBefore(range.getRow());
};

const insertRowBelow: RunFn = (api) => {
  const range = activeRange(api);
  const sheet = activeSheet(api);
  if (!range || !sheet) return;
  sheet.insertRowAfter(range.getRow() + range.getHeight() - 1);
};

const insertColumnLeft: RunFn = (api) => {
  const range = activeRange(api);
  const sheet = activeSheet(api);
  if (!range || !sheet) return;
  sheet.insertColumnBefore(range.getColumn());
};

const insertColumnRight: RunFn = (api) => {
  const range = activeRange(api);
  const sheet = activeSheet(api);
  if (!range || !sheet) return;
  sheet.insertColumnAfter(range.getColumn() + range.getWidth() - 1);
};

const deleteSelectedRow: RunFn = (api) => {
  const range = activeRange(api);
  const sheet = activeSheet(api);
  if (!range || !sheet) return;
  sheet.deleteRows(range.getRow(), range.getHeight());
};

const deleteSelectedColumn: RunFn = (api) => {
  const range = activeRange(api);
  const sheet = activeSheet(api);
  if (!range || !sheet) return;
  sheet.deleteColumns(range.getColumn(), range.getWidth());
};

const insertNewSheet: RunFn = (api) => {
  fu(api).getActiveWorkbook()?.insertSheet();
};

const insertImage: RunFn = (api) => void api.executeCommand('sheet.command.insert-float-image');

const insertHyperlink: RunFn = (api) =>
  void api.executeCommand('sheet.operation.insert-hyper-link');

const insertComment: RunFn = (api) => void api.executeCommand('sheet.operation.show-comment-modal');

const insertTodayDate: RunFn = (api) => {
  const range = activeRange(api);
  if (!range) return;
  const today = new Date();
  const pad = (n: number) => n.toString().padStart(2, '0');
  const v = `${today.getFullYear()}-${pad(today.getMonth() + 1)}-${pad(today.getDate())}`;
  range.setValue({ v });
};

const insertCurrentTime: RunFn = (api) => {
  const range = activeRange(api);
  if (!range) return;
  const now = new Date();
  const pad = (n: number) => n.toString().padStart(2, '0');
  const v = `${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`;
  range.setValue({ v });
};

const insertTable: RunFn = (api) => {
  void (async () => {
    const range = activeRange(api);
    const sheet = activeSheet(api);
    const wb = fu(api).getActiveWorkbook();
    if (!range || !sheet || !wb) return;
    await ensurePluginByName('table');
    await api.executeCommand('sheet.command.add-table', {
      unitId: wb.getId(),
      subUnitId: sheet.getSheetId(),
      range: {
        startRow: range.getRow(),
        startColumn: range.getColumn(),
        endRow: range.getRow() + range.getHeight() - 1,
        endColumn: range.getColumn() + range.getWidth() - 1,
      },
    });
  })();
};

const autoFitColumns: RunFn = (api) => {
  const range = activeRange(api);
  const sheet = activeSheet(api);
  if (!range || !sheet) return;
  const withAutoWidth = sheet as unknown as {
    setColumnAutoWidth?: (col: number, n: number) => unknown;
  };
  withAutoWidth.setColumnAutoWidth?.(range.getColumn(), range.getWidth());
};

const AUTO_FIT_ROW_CAP = 500;
const autoFitRows: RunFn = (api) => {
  const range = activeRange(api);
  const sheet = activeSheet(api);
  if (!range || !sheet) return;
  const start = range.getRow();
  const count = Math.min(range.getHeight(), AUTO_FIT_ROW_CAP);
  for (let r = 0; r < count; r++) sheet.autoFitRow(start + r);
};

/* ── Format ───────────────────────────────────────────────────────────── */

// Patterns copied from home-tab-actions.ts NUMBER_FORMAT_PATTERNS.
const NUMBER_FORMAT_PATTERNS = {
  general: '',
  number: '#,##0.00',
  integer: '#,##0',
  currency: '"$"#,##0.00',
  accounting: '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)',
  percent: '0.00%',
  date: 'yyyy-mm-dd',
  time: 'hh:mm:ss',
  scientific: '0.00E+00',
  text: '@',
} as const;
type NumberFormatKey = keyof typeof NUMBER_FORMAT_PATTERNS;

const NUM_FORMAT_ORDER: NumberFormatKey[] = [
  'general',
  'number',
  'integer',
  'currency',
  'accounting',
  'percent',
  'date',
  'time',
  'scientific',
  'text',
];

const NUM_FORMAT_SHORTCUT: Partial<Record<NumberFormatKey, string>> = {
  number: 'Ctrl+Shift+1',
  time: 'Ctrl+Shift+2',
  date: 'Ctrl+Shift+3',
  currency: 'Ctrl+Shift+4',
  percent: 'Ctrl+Shift+5',
  scientific: 'Ctrl+Shift+6',
};

function setNumberFormatByKey(api: CasualSheetsAPI, key: NumberFormatKey) {
  // setNumberFormat lives on the sheets-numfmt facade extension (runtime cast).
  const range = activeRange(api) as unknown as { setNumberFormat?: (p: string) => unknown } | null;
  range?.setNumberFormat?.(NUMBER_FORMAT_PATTERNS[key]);
}

const increaseDecimal: RunFn = (api) =>
  void api.executeCommand('sheet.command.numfmt.add.decimal.command');
const decreaseDecimal: RunFn = (api) =>
  void api.executeCommand('sheet.command.numfmt.subtract.decimal.command');
const clearFormat: RunFn = (api) => void api.executeCommand('sheet.command.clear-selection-format');

function applyBorders(api: CasualSheetsAPI, choice: 'all' | 'outside' | 'none') {
  const range = activeRange(api);
  if (!range) return;
  const type =
    choice === 'all' ? BorderType.ALL : choice === 'outside' ? BorderType.OUTSIDE : BorderType.NONE;
  const style = choice === 'none' ? BorderStyleTypes.NONE : BorderStyleTypes.THIN;
  range.setBorder(type, style, '#000000');
}

/* ── Format → Visibility (command-driven; mirror tab-actions.ts) ───────── */

function rowSpan(api: CasualSheetsAPI) {
  const wb = fu(api).getActiveWorkbook();
  const sheet = activeSheet(api);
  const range = activeRange(api);
  if (!wb || !sheet || !range) return null;
  const startRow = range.getRow();
  const endRow = startRow + range.getHeight() - 1;
  const maxCol = (sheet as unknown as { getMaxColumns?: () => number }).getMaxColumns?.() ?? 1;
  return { wb, sheet, startRow, endRow, maxCol };
}

function colSpan(api: CasualSheetsAPI) {
  const wb = fu(api).getActiveWorkbook();
  const sheet = activeSheet(api);
  const range = activeRange(api);
  if (!wb || !sheet || !range) return null;
  const startColumn = range.getColumn();
  const endColumn = startColumn + range.getWidth() - 1;
  const maxRow = (sheet as unknown as { getMaxRows?: () => number }).getMaxRows?.() ?? 1;
  return { wb, sheet, startColumn, endColumn, maxRow };
}

const hideSelectedRows: RunFn = (api) => {
  const s = rowSpan(api);
  if (!s) return;
  void api.executeCommand('sheet.command.set-rows-hidden', {
    unitId: s.wb.getId(),
    subUnitId: s.sheet.getSheetId(),
    ranges: [
      {
        startRow: s.startRow,
        endRow: s.endRow,
        startColumn: 0,
        endColumn: Math.max(0, s.maxCol - 1),
        rangeType: 1,
      },
    ],
  });
};

const unhideSelectedRows: RunFn = (api) => {
  const s = rowSpan(api);
  if (!s) return;
  void api.executeCommand('sheet.command.set-specific-rows-visible', {
    unitId: s.wb.getId(),
    subUnitId: s.sheet.getSheetId(),
    ranges: [
      {
        startRow: s.startRow,
        endRow: s.endRow,
        startColumn: 0,
        endColumn: Math.max(0, s.maxCol - 1),
        rangeType: 1,
      },
    ],
  });
};

const hideSelectedColumns: RunFn = (api) => {
  const s = colSpan(api);
  if (!s) return;
  void api.executeCommand('sheet.command.set-col-hidden', {
    unitId: s.wb.getId(),
    subUnitId: s.sheet.getSheetId(),
    ranges: [
      {
        startRow: 0,
        endRow: Math.max(0, s.maxRow - 1),
        startColumn: s.startColumn,
        endColumn: s.endColumn,
        rangeType: 2,
      },
    ],
  });
};

const unhideSelectedColumns: RunFn = (api) => {
  const s = colSpan(api);
  if (!s) return;
  // Univer's "show specific cols" id is asymmetric with the row variant.
  void api.executeCommand('sheet.command.set-col-visible-on-cols', {
    unitId: s.wb.getId(),
    subUnitId: s.sheet.getSheetId(),
    ranges: [
      {
        startRow: 0,
        endRow: Math.max(0, s.maxRow - 1),
        startColumn: s.startColumn,
        endColumn: s.endColumn,
        rangeType: 2,
      },
    ],
  });
};

/* ── View ─────────────────────────────────────────────────────────────── */

type FreezeCapableSheet = {
  setFrozenRows: (n: number) => unknown;
  setFrozenColumns: (n: number) => unknown;
};

const freezeFirstRow: RunFn = (api) => {
  (activeSheet(api) as unknown as FreezeCapableSheet | null)?.setFrozenRows(1);
};
const freezeFirstColumn: RunFn = (api) => {
  (activeSheet(api) as unknown as FreezeCapableSheet | null)?.setFrozenColumns(1);
};
const freezeAtSelection: RunFn = (api) =>
  void api.executeCommand('sheet.command.set-selection-frozen');
const unfreezePanes: RunFn = (api) => void api.executeCommand('sheet.command.cancel-frozen');

const toggleGridlines: RunFn = (api) => {
  const wb = fu(api).getActiveWorkbook();
  const sheet = activeSheet(api);
  if (!wb || !sheet) return;
  // BooleanNumber: 0 = hide, 1 = show. The app reads current state; without a
  // reactive UI store here we hide (matches the app's default click path).
  void api.executeCommand('sheet.command.toggle-gridlines', {
    unitId: wb.getId(),
    subUnitId: sheet.getSheetId(),
    showGridlines: 0,
  });
};

const toggleCommentPanel: RunFn = (api) =>
  void api.executeCommand('sheet.operation.toggle-comment-panel');

const jumpToFirstCell: RunFn = (api) => {
  activeSheet(api)?.getRange(0, 0).activate();
};

const switchToPreviousSheet: RunFn = (api) => switchSheetByDelta(api, -1);
const switchToNextSheet: RunFn = (api) => switchSheetByDelta(api, +1);

function switchSheetByDelta(api: CasualSheetsAPI, delta: -1 | 1) {
  const wb = fu(api).getActiveWorkbook();
  const active = wb?.getActiveSheet();
  if (!wb || !active) return;
  const sheets = wb.getSheets();
  const activeId = active.getSheetId();
  const idx = sheets.findIndex((s) => s.getSheetId() === activeId);
  if (idx < 0) return;
  const nextIdx = idx + delta;
  if (nextIdx < 0 || nextIdx >= sheets.length) return;
  wb.setActiveSheet(sheets[nextIdx]);
}

const showFormulas: RunFn = (api) => void api.executeCommand('sheet.command.set-show-formula', {});

/* ── Data ─────────────────────────────────────────────────────────────── */

const sortAsc: RunFn = (api) => sortRange(api, true);
const sortDesc: RunFn = (api) => sortRange(api, false);

function sortRange(api: CasualSheetsAPI, ascending: boolean) {
  const range = activeRange(api);
  if (!range) return;
  // sort() comes from the sheets-sort facade extension (runtime cast).
  const withSort = range as unknown as {
    sort?: (spec: { column: number; ascending: boolean }) => unknown;
  };
  withSort.sort?.({ column: range.getColumn(), ascending });
}

const toggleFilter: RunFn = (api) => {
  void (async () => {
    await ensurePluginByName('filter');
    const wb = fu(api).getActiveWorkbook();
    const sheet = activeSheet(api);
    const range = activeRange(api);
    if (!wb || !sheet || !range) return;
    const sheetWithFilter = sheet as unknown as { getFilter?: () => unknown };
    if (sheetWithFilter.getFilter?.()) {
      await api.executeCommand('sheet.command.remove-sheet-filter', {
        unitId: wb.getId(),
        subUnitId: sheet.getSheetId(),
      });
      return;
    }
    await api.executeCommand('sheet.command.set-filter-range', {
      unitId: wb.getId(),
      subUnitId: sheet.getSheetId(),
      range: {
        startRow: range.getRow(),
        startColumn: range.getColumn(),
        endRow: range.getRow() + range.getHeight() - 1,
        endColumn: range.getColumn() + range.getWidth() - 1,
      },
    });
  })();
};

const splitTextToColumns: RunFn = (api) =>
  void api.executeCommand('sheet.command.split-text-to-columns');

const forceRecalculate: RunFn = (api) =>
  void api.executeCommand('formula.mutation.set-formula-calculation-start', {
    forceCalculation: true,
  });

/* ───────────────────────────── menu structure ─────────────────────────── */

// Mirrors the app's menu order. Dialog items resolve `dialog` against
// `onDialogRequest`; non-dialog items dispatch via the inline helpers above.
const MENUS: MenuDef[] = [
  {
    id: 'file',
    label: 'File',
    feature: 'file',
    items: [
      { kind: 'item', id: 'properties', label: 'Properties…', icon: 'info', dialog: 'properties' },
      {
        kind: 'item',
        id: 'about',
        label: 'About casual sheets',
        icon: 'help_outline',
        dialog: 'about',
        feature: 'branding',
      },
    ],
  },
  {
    id: 'edit',
    label: 'Edit',
    items: [
      { kind: 'item', id: 'undo', label: 'Undo', icon: 'undo', shortcut: 'Ctrl+Z', run: undo },
      { kind: 'item', id: 'redo', label: 'Redo', icon: 'redo', shortcut: 'Ctrl+Y', run: redo },
      { kind: 'separator', id: 'sep-clip' },
      { kind: 'item', id: 'cut', label: 'Cut', icon: 'content_cut', shortcut: 'Ctrl+X', run: cut },
      {
        kind: 'item',
        id: 'copy',
        label: 'Copy',
        icon: 'content_copy',
        shortcut: 'Ctrl+C',
        run: copy,
      },
      {
        kind: 'item',
        id: 'paste',
        label: 'Paste',
        icon: 'content_paste',
        shortcut: 'Ctrl+V',
        run: paste,
      },
      {
        kind: 'item',
        id: 'paste-format',
        label: 'Paste formatting only',
        icon: 'content_paste',
        shortcut: 'Ctrl+Shift+V',
        run: pasteFormattingOnly,
      },
      {
        kind: 'item',
        id: 'paste-special',
        label: 'Paste Special…',
        icon: 'content_paste_go',
        shortcut: 'Ctrl+Alt+V',
        dialog: 'paste-special',
      },
      { kind: 'separator', id: 'sep-find' },
      {
        kind: 'item',
        id: 'find-replace',
        label: 'Find & Replace…',
        icon: 'search',
        shortcut: 'Ctrl+F',
        dialog: 'find-replace',
      },
      { kind: 'separator', id: 'sep-cells' },
      {
        kind: 'item',
        id: 'edit-insert-cells',
        label: 'Insert cells…',
        icon: 'add_box',
        shortcut: 'Ctrl++',
        dialog: 'insert-cells',
      },
      {
        kind: 'item',
        id: 'edit-delete-cells',
        label: 'Delete cells…',
        icon: 'indeterminate_check_box',
        shortcut: 'Ctrl+-',
        dialog: 'delete-cells',
      },
    ],
  },
  {
    id: 'view',
    label: 'View',
    items: [
      {
        kind: 'item',
        id: 'show-formulas',
        label: 'Show formulas',
        icon: 'description',
        shortcut: 'Ctrl+`',
        run: showFormulas,
      },
      {
        kind: 'item',
        id: 'toggle-gridlines',
        label: 'Toggle gridlines',
        icon: 'grid_on',
        run: toggleGridlines,
      },
      { kind: 'separator', id: 'sep-freeze' },
      {
        kind: 'item',
        id: 'freeze-row',
        label: 'Freeze top row',
        icon: 'border_horizontal',
        run: freezeFirstRow,
      },
      {
        kind: 'item',
        id: 'freeze-col',
        label: 'Freeze first column',
        icon: 'border_vertical',
        run: freezeFirstColumn,
      },
      {
        kind: 'item',
        id: 'freeze-selection',
        label: 'Freeze panes (at selection)',
        icon: 'grid_4x4',
        run: freezeAtSelection,
      },
      { kind: 'item', id: 'unfreeze', label: 'Unfreeze', icon: 'grid_off', run: unfreezePanes },
      { kind: 'separator', id: 'sep-nav' },
      {
        kind: 'item',
        id: 'jump-home',
        label: 'Jump to A1',
        icon: 'home',
        shortcut: 'Ctrl+Home',
        run: jumpToFirstCell,
      },
      {
        kind: 'item',
        id: 'prev-sheet',
        label: 'Previous sheet',
        icon: 'navigate_before',
        shortcut: 'Ctrl+PageUp',
        run: switchToPreviousSheet,
      },
      {
        kind: 'item',
        id: 'next-sheet',
        label: 'Next sheet',
        icon: 'navigate_next',
        shortcut: 'Ctrl+PageDown',
        run: switchToNextSheet,
      },
      { kind: 'separator', id: 'sep-panels' },
      {
        kind: 'item',
        id: 'comments-panel',
        label: 'Comments panel',
        icon: 'forum',
        run: toggleCommentPanel,
      },
    ],
  },
  {
    id: 'insert',
    label: 'Insert',
    items: [
      {
        kind: 'item',
        id: 'new-sheet',
        label: 'New sheet',
        icon: 'add_box',
        shortcut: 'Shift+F11',
        run: insertNewSheet,
      },
      {
        kind: 'item',
        id: 'insert-table',
        label: 'Table',
        icon: 'table_rows',
        shortcut: 'Ctrl+L',
        run: insertTable,
        feature: 'tables',
      },
      {
        kind: 'item',
        id: 'insert-chart',
        label: 'Chart…',
        icon: 'bar_chart',
        dialog: 'insert-chart',
        feature: 'charts',
      },
      {
        kind: 'item',
        id: 'insert-sparkline',
        label: 'Sparkline…',
        icon: 'show_chart',
        dialog: 'insert-sparkline',
        feature: 'sparklines',
      },
      {
        kind: 'item',
        id: 'insert-pivot',
        label: 'PivotTable…',
        icon: 'pivot_table_chart',
        dialog: 'insert-pivot',
        feature: 'pivots',
      },
      { kind: 'separator', id: 'sep-objects' },
      { kind: 'item', id: 'insert-image', label: 'Image…', icon: 'image', run: insertImage },
      {
        kind: 'item',
        id: 'insert-function',
        label: 'Function…',
        icon: 'functions',
        shortcut: 'Shift+F3',
        dialog: 'insert-function',
      },
      {
        kind: 'item',
        id: 'insert-link',
        label: 'Hyperlink…',
        icon: 'link',
        shortcut: 'Ctrl+K',
        run: insertHyperlink,
      },
      {
        kind: 'item',
        id: 'insert-comment',
        label: 'Comment',
        icon: 'comment',
        shortcut: 'Shift+F2',
        run: insertComment,
      },
      { kind: 'separator', id: 'sep-rowcol' },
      {
        kind: 'submenu',
        id: 'insert-rowcol',
        label: 'Rows & columns',
        icon: 'grid_on',
        items: [
          {
            kind: 'item',
            id: 'insert-row-above',
            label: 'Row above',
            icon: 'vertical_align_top',
            run: insertRowAbove,
          },
          {
            kind: 'item',
            id: 'insert-row-below',
            label: 'Row below',
            icon: 'vertical_align_bottom',
            run: insertRowBelow,
          },
          {
            kind: 'item',
            id: 'insert-col-left',
            label: 'Column left',
            icon: 'keyboard_tab_rtl',
            run: insertColumnLeft,
          },
          {
            kind: 'item',
            id: 'insert-col-right',
            label: 'Column right',
            icon: 'keyboard_tab',
            run: insertColumnRight,
          },
        ],
      },
      { kind: 'separator', id: 'sep-autofit' },
      {
        kind: 'item',
        id: 'autofit-col',
        label: 'Auto-fit column width',
        icon: 'settings_ethernet',
        run: autoFitColumns,
      },
      {
        kind: 'item',
        id: 'autofit-row',
        label: 'Auto-fit row height',
        icon: 'height',
        run: autoFitRows,
      },
      { kind: 'separator', id: 'sep-date' },
      {
        kind: 'item',
        id: 'insert-today',
        label: "Today's date",
        icon: 'today',
        shortcut: 'Ctrl+;',
        run: insertTodayDate,
      },
      {
        kind: 'item',
        id: 'insert-time',
        label: 'Current time',
        icon: 'schedule',
        shortcut: 'Ctrl+Shift+:',
        run: insertCurrentTime,
      },
    ],
  },
  {
    id: 'format',
    label: 'Format',
    items: [
      {
        kind: 'item',
        id: 'format-cells',
        label: 'Format cells…',
        icon: 'format_shapes',
        shortcut: 'Ctrl+1',
        dialog: 'format-cells',
      },
      { kind: 'separator', id: 'sep-format-cells' },
      {
        kind: 'item',
        id: 'bold',
        label: 'Bold',
        icon: 'format_bold',
        shortcut: 'Ctrl+B',
        run: (api) => void api.executeCommand('sheet.command.set-range-bold'),
      },
      {
        kind: 'item',
        id: 'italic',
        label: 'Italic',
        icon: 'format_italic',
        shortcut: 'Ctrl+I',
        run: (api) => void api.executeCommand('sheet.command.set-range-italic'),
      },
      {
        kind: 'item',
        id: 'underline',
        label: 'Underline',
        icon: 'format_underlined',
        shortcut: 'Ctrl+U',
        run: (api) => void api.executeCommand('sheet.command.set-range-underline'),
      },
      {
        kind: 'item',
        id: 'wrap-text',
        label: 'Wrap text',
        icon: 'wrap_text',
        run: (api) => void api.executeCommand('sheet.command.set-text-wrap', { value: 3 }),
      },
      { kind: 'separator', id: 'sep-numfmt' },
      {
        kind: 'submenu',
        id: 'num-format',
        label: 'Number format',
        icon: 'looks_one',
        items: NUM_FORMAT_ORDER.map<MenuItemDef>((k) => ({
          kind: 'item',
          id: `num-${k}`,
          label: k[0]!.toUpperCase() + k.slice(1),
          icon: 'looks_one',
          shortcut: NUM_FORMAT_SHORTCUT[k],
          run: (api) => setNumberFormatByKey(api, k),
        })),
      },
      {
        kind: 'item',
        id: 'decimal-up',
        label: 'Increase decimals',
        icon: 'decimal_increase',
        run: increaseDecimal,
      },
      {
        kind: 'item',
        id: 'decimal-down',
        label: 'Decrease decimals',
        icon: 'decimal_decrease',
        run: decreaseDecimal,
      },
      { kind: 'separator', id: 'sep-borders' },
      {
        kind: 'submenu',
        id: 'borders',
        label: 'Borders',
        icon: 'border_all',
        items: [
          {
            kind: 'item',
            id: 'border-all',
            label: 'All borders',
            icon: 'border_all',
            run: (api) => applyBorders(api, 'all'),
          },
          {
            kind: 'item',
            id: 'border-outside',
            label: 'Outside borders',
            icon: 'border_outer',
            run: (api) => applyBorders(api, 'outside'),
          },
          {
            kind: 'item',
            id: 'border-none',
            label: 'No border',
            icon: 'border_clear',
            run: (api) => applyBorders(api, 'none'),
          },
        ],
      },
      { kind: 'separator', id: 'sep-cond' },
      {
        kind: 'item',
        id: 'conditional-formatting',
        label: 'Conditional formatting…',
        icon: 'palette',
        dialog: 'conditional-formatting',
        feature: 'conditionalFormatting',
      },
      { kind: 'separator', id: 'sep-visibility' },
      {
        kind: 'submenu',
        id: 'visibility',
        label: 'Visibility',
        icon: 'visibility',
        items: [
          {
            kind: 'item',
            id: 'hide-row',
            label: 'Hide row',
            icon: 'visibility_off',
            shortcut: 'Ctrl+9',
            run: hideSelectedRows,
          },
          {
            kind: 'item',
            id: 'unhide-row',
            label: 'Unhide row',
            icon: 'visibility',
            shortcut: 'Ctrl+Shift+9',
            run: unhideSelectedRows,
          },
          {
            kind: 'item',
            id: 'hide-col',
            label: 'Hide column',
            icon: 'visibility_off',
            shortcut: 'Ctrl+0',
            run: hideSelectedColumns,
          },
          {
            kind: 'item',
            id: 'unhide-col',
            label: 'Unhide column',
            icon: 'visibility',
            shortcut: 'Ctrl+Shift+0',
            run: unhideSelectedColumns,
          },
        ],
      },
      { kind: 'separator', id: 'sep-clear' },
      {
        kind: 'item',
        id: 'clear-format',
        label: 'Clear formatting',
        icon: 'format_clear',
        run: clearFormat,
      },
      { kind: 'separator', id: 'sep-delete' },
      {
        kind: 'item',
        id: 'delete-row',
        label: 'Delete row',
        icon: 'delete_sweep',
        run: deleteSelectedRow,
      },
      {
        kind: 'item',
        id: 'delete-col',
        label: 'Delete column',
        icon: 'folder_delete',
        run: deleteSelectedColumn,
      },
    ],
  },
  {
    id: 'data',
    label: 'Data',
    items: [
      {
        kind: 'item',
        id: 'sort-asc',
        label: 'Sort ascending',
        icon: 'arrow_upward',
        run: sortAsc,
      },
      {
        kind: 'item',
        id: 'sort-desc',
        label: 'Sort descending',
        icon: 'arrow_downward',
        run: sortDesc,
      },
      {
        kind: 'item',
        id: 'sort-custom',
        label: 'Sort range…',
        icon: 'sort',
        dialog: 'custom-sort',
      },
      {
        kind: 'item',
        id: 'toggle-filter',
        label: 'Toggle filter',
        icon: 'filter_alt',
        shortcut: 'Ctrl+Shift+L',
        run: toggleFilter,
        feature: 'filter',
      },
      { kind: 'separator', id: 'sep-tools' },
      {
        kind: 'item',
        id: 'data-validation',
        label: 'Data validation…',
        icon: 'rule',
        dialog: 'data-validation',
        feature: 'dataValidation',
      },
      {
        kind: 'item',
        id: 'name-manager',
        label: 'Name Manager…',
        icon: 'bookmark_add',
        shortcut: 'Ctrl+F3',
        dialog: 'name-manager',
      },
      {
        kind: 'item',
        id: 'goal-seek',
        label: 'Goal Seek…',
        icon: 'analytics',
        dialog: 'goal-seek',
      },
      { kind: 'separator', id: 'sep-clean' },
      {
        kind: 'item',
        id: 'text-to-columns',
        label: 'Text to Columns',
        icon: 'splitscreen',
        run: splitTextToColumns,
      },
      {
        kind: 'item',
        id: 'recalculate',
        label: 'Recalculate',
        icon: 'autorenew',
        shortcut: 'F9',
        run: forceRecalculate,
      },
    ],
  },
  {
    id: 'help',
    label: 'Help',
    // Whole-menu gate: an embedded host passes `features={{ help: false }}` to
    // drop the Help menu entirely (no editor-branded surface). Individual
    // branding links below also carry `feature: 'branding'` so a host that
    // keeps Help (for keyboard shortcuts) can still hide the About / GitHub
    // links with `features={{ branding: false }}`.
    feature: 'help',
    items: [
      {
        kind: 'item',
        id: 'keyboard-shortcuts',
        label: 'Keyboard shortcuts',
        icon: 'info',
        shortcut: 'Ctrl+/',
        dialog: 'keyboard-shortcuts',
      },
      { kind: 'separator', id: 'sep-help', feature: 'branding' },
      {
        kind: 'item',
        id: 'about',
        label: 'About casual sheets',
        icon: 'info',
        dialog: 'about',
        feature: 'branding',
      },
      {
        kind: 'item',
        id: 'github',
        label: 'View on GitHub',
        icon: 'open_in_new',
        run: () => openExternal('https://github.com/CasualOffice/sheets'),
        feature: 'branding',
      },
    ],
  },
];

/* ───────────────────────────── styling ────────────────────────────────── */

const BAR_STYLE: CSSProperties = {
  position: 'relative',
  display: 'flex',
  alignItems: 'center',
  gap: 2,
  padding: '2px 6px',
  borderBottom: '1px solid var(--cs-chrome-border, #e6e9ee)',
  background: 'var(--cs-chrome-bg, #eef1f5)',
  flex: '0 0 auto',
  userSelect: 'none',
  font: 'inherit',
  fontSize: 13,
};

const MENU_BTN_STYLE: CSSProperties = {
  height: 26,
  display: 'inline-flex',
  alignItems: 'center',
  padding: '0 10px',
  border: 'none',
  borderRadius: 6,
  background: 'transparent',
  cursor: 'pointer',
  color: 'var(--cs-chrome-fg, #201f1e)',
  font: 'inherit',
  fontSize: 13,
};

const DROPDOWN_STYLE: CSSProperties = {
  position: 'absolute',
  top: '100%',
  minWidth: 220,
  marginTop: 2,
  padding: 4,
  display: 'flex',
  flexDirection: 'column',
  border: '1px solid var(--cs-chrome-border, #e6e9ee)',
  borderRadius: 8,
  background: 'var(--cs-chrome-input-bg, #fff)',
  boxShadow: '0 6px 20px rgba(0,0,0,0.16)',
  zIndex: 1000,
};

const ITEM_STYLE: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 10,
  width: '100%',
  height: 30,
  padding: '0 10px',
  border: 'none',
  borderRadius: 6,
  background: 'transparent',
  cursor: 'pointer',
  color: 'var(--cs-chrome-fg, #201f1e)',
  font: 'inherit',
  fontSize: 13,
  textAlign: 'left',
};

const SEPARATOR_STYLE: CSSProperties = {
  height: 1,
  margin: '4px 6px',
  background: 'var(--cs-chrome-border, #e6e9ee)',
};

const SHORTCUT_STYLE: CSSProperties = {
  marginLeft: 'auto',
  paddingLeft: 16,
  fontSize: 11,
  color: 'var(--cs-chrome-muted, #6b7280)',
};

const SUBMENU_PANEL_STYLE: CSSProperties = {
  position: 'absolute',
  top: -4,
  left: '100%',
  minWidth: 200,
  marginLeft: 2,
  padding: 4,
  display: 'flex',
  flexDirection: 'column',
  border: '1px solid var(--cs-chrome-border, #e6e9ee)',
  borderRadius: 8,
  background: 'var(--cs-chrome-input-bg, #fff)',
  boxShadow: '0 6px 20px rgba(0,0,0,0.16)',
  zIndex: 1001,
};

/* ───────────────────────────── component ──────────────────────────────── */

export interface MenuBarProps {
  /** Live API, or `null` until the editor is ready. */
  api: CasualSheetsAPI | null;
  /**
   * Per-feature toggles. A control / menu whose `feature` key is `false` is not
   * rendered. Unknown / omitted keys default to enabled.
   */
  features?: Record<string, boolean>;
  /**
   * Host chrome extensions — custom menu items appended to their target menu.
   * Dialogs/toolbar/panels are handled elsewhere; only `extensions.menu` is read
   * here.
   */
  extensions?: ChromeExtensions;
}

export function MenuBar({ api, features = {}, extensions }: MenuBarProps) {
  const dialogs = useDialogs();
  const [open, setOpen] = useState<MenuId | null>(null);
  const [openSubmenu, setOpenSubmenu] = useState<string | null>(null);
  const rootRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    ensureChromeFonts();
  }, []);

  // Close on Escape + on a pointerdown outside the menu bar.
  useEffect(() => {
    if (open === null) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        setOpen(null);
        setOpenSubmenu(null);
      }
    };
    const onDown = (e: PointerEvent) => {
      if (!rootRef.current?.contains(e.target as Node)) {
        setOpen(null);
        setOpenSubmenu(null);
      }
    };
    document.addEventListener('keydown', onKey);
    document.addEventListener('pointerdown', onDown, true);
    return () => {
      document.removeEventListener('keydown', onKey);
      document.removeEventListener('pointerdown', onDown, true);
    };
  }, [open]);

  const close = () => {
    setOpen(null);
    setOpenSubmenu(null);
  };

  const runItem = (item: Extract<MenuItemDef, { kind: 'item' }>) => {
    close();
    if (!api) return;
    if (item.dialog) {
      // Range-seeded dialogs get the current selection in A1 as context.
      const seeded =
        item.dialog === 'insert-chart' ||
        item.dialog === 'insert-pivot' ||
        item.dialog === 'insert-sparkline' ||
        item.dialog === 'insert-cells' ||
        item.dialog === 'delete-cells';
      // openDialog opens the SDK built-in by default; a host override / host-owned
      // kind / onDialogRequest wins per the resolution rules in DialogProvider.
      dialogs.openDialog(item.dialog, seeded ? selectionA1(api) : undefined);
      return;
    }
    item.run?.(api);
  };

  // Compute the visible menus once per render (feature + dialog-open gating +
  // host menu extensions). Extracted to `computeVisibleMenus` so the gating
  // contract is unit-testable without a DOM.
  const visibleMenus = computeVisibleMenus(MENUS, features, dialogs.canOpen, extensions?.menu);

  const renderItems = (items: MenuItemDef[]): ReactNode =>
    items.map((item) => {
      if (item.kind === 'separator') {
        return <div key={item.id} style={SEPARATOR_STYLE} role="separator" aria-hidden />;
      }
      if (item.kind === 'submenu') {
        const isSubOpen = openSubmenu === item.id;
        return (
          <div
            key={item.id}
            style={{ position: 'relative' }}
            onMouseEnter={() => setOpenSubmenu(item.id)}
            onMouseLeave={() => setOpenSubmenu((cur) => (cur === item.id ? null : cur))}
          >
            <button
              type="button"
              role="menuitem"
              data-testid={`cs-menuitem-${item.id}`}
              aria-haspopup="menu"
              aria-expanded={isSubOpen}
              disabled={!api}
              style={{ ...ITEM_STYLE, opacity: api ? 1 : 0.5 }}
              onMouseDown={(e) => e.preventDefault()}
            >
              {item.icon ? (
                <Icon name={item.icon} size={18} />
              ) : (
                <span style={{ width: 18 }} aria-hidden />
              )}
              <span>{item.label}</span>
              <Icon name="chevron_right" size={18} style={{ marginLeft: 'auto' }} />
            </button>
            {isSubOpen && (
              <div style={SUBMENU_PANEL_STYLE} role="menu" aria-label={item.label}>
                {renderItems(item.items)}
              </div>
            )}
          </div>
        );
      }
      return (
        <button
          key={item.id}
          type="button"
          role="menuitem"
          data-testid={`cs-menuitem-${item.id}`}
          disabled={!api}
          style={{ ...ITEM_STYLE, opacity: api ? 1 : 0.5 }}
          onMouseDown={(e) => {
            e.preventDefault();
            runItem(item);
          }}
          onMouseEnter={(e) => {
            setOpenSubmenu(null);
            e.currentTarget.style.background = 'var(--cs-chrome-hover, rgba(0,0,0,0.06))';
          }}
          onMouseLeave={(e) => {
            e.currentTarget.style.background = 'transparent';
          }}
        >
          {item.icon ? (
            <Icon name={item.icon} size={18} />
          ) : (
            <span style={{ width: 18 }} aria-hidden />
          )}
          <span>{item.label}</span>
          {item.shortcut && <span style={SHORTCUT_STYLE}>{fmtShortcut(item.shortcut)}</span>}
        </button>
      );
    });

  return (
    <div
      ref={rootRef}
      style={BAR_STYLE}
      data-testid="cs-menubar"
      role="menubar"
      aria-label="Menu bar"
    >
      {visibleMenus.map((menu) => {
        const isOpen = open === menu.id;
        return (
          <div key={menu.id} style={{ position: 'relative' }}>
            <button
              type="button"
              data-menu={menu.id}
              data-testid={`cs-menu-${menu.id}`}
              aria-haspopup="menu"
              aria-expanded={isOpen}
              style={{
                ...MENU_BTN_STYLE,
                background: isOpen ? 'var(--cs-chrome-active, #e6f3f7)' : 'transparent',
                color: isOpen ? 'var(--cs-chrome-active-fg, #0e7490)' : MENU_BTN_STYLE.color,
              }}
              // mousedown (not click) so the grid selection isn't lost first;
              // toggle the menu open/closed.
              onMouseDown={(e) => {
                e.preventDefault();
                setOpenSubmenu(null);
                setOpen((cur) => (cur === menu.id ? null : menu.id));
              }}
              onMouseEnter={(e) => {
                // Hover-to-switch once a menu is already open (Office behaviour).
                if (open !== null && open !== menu.id) {
                  setOpen(menu.id);
                  setOpenSubmenu(null);
                } else if (!isOpen)
                  e.currentTarget.style.background = 'var(--cs-chrome-hover, rgba(0,0,0,0.06))';
              }}
              onMouseLeave={(e) => {
                e.currentTarget.style.background = isOpen
                  ? 'var(--cs-chrome-active, #e6f3f7)'
                  : 'transparent';
              }}
            >
              {menu.label}
            </button>
            {isOpen && (
              <div style={DROPDOWN_STYLE} role="menu" aria-label={menu.label}>
                {renderItems(menu.items)}
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}
