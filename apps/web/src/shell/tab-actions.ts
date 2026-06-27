import type { FUniver } from '@univerjs/core/facade';
import type { FRange } from '@univerjs/sheets/facade';
import { ensurePluginByName } from '@casualoffice/sheets/univer';
import {
  activateRange,
  activeRange,
  activeSheet,
  isHidden,
  maxColumns,
  maxRows,
  rangeBox as facadeRangeBox,
  saveWorkbook,
  sheetId as facadeSheetId,
} from '../univer-facade';

/** Imperative dispatchers used by Insert / Formulas / Data tabs. */

/* ── Insert tab ─────────────────────────────────────────────────────────── */

export function insertRowAbove(api: FUniver) {
  const range = activeRange(api);
  const sheet = activeSheet(api);
  if (!range || !sheet) return;
  sheet.insertRowBefore(range.getRow());
}

export function insertRowBelow(api: FUniver) {
  const range = activeRange(api);
  const sheet = activeSheet(api);
  if (!range || !sheet) return;
  sheet.insertRowAfter(range.getRow() + range.getHeight() - 1);
}

export function insertColumnLeft(api: FUniver) {
  const range = activeRange(api);
  const sheet = activeSheet(api);
  if (!range || !sheet) return;
  sheet.insertColumnBefore(range.getColumn());
}

export function insertColumnRight(api: FUniver) {
  const range = activeRange(api);
  const sheet = activeSheet(api);
  if (!range || !sheet) return;
  sheet.insertColumnAfter(range.getColumn() + range.getWidth() - 1);
}

export function deleteSelectedRow(api: FUniver) {
  const range = activeRange(api);
  const sheet = activeSheet(api);
  if (!range || !sheet) return;
  sheet.deleteRows(range.getRow(), range.getHeight());
}

export function deleteSelectedColumn(api: FUniver) {
  const range = activeRange(api);
  const sheet = activeSheet(api);
  if (!range || !sheet) return;
  sheet.deleteColumns(range.getColumn(), range.getWidth());
}

export function insertNewSheet(api: FUniver) {
  const wb = api.getActiveWorkbook();
  if (!wb) return;
  wb.insertSheet();
}

/** Excel's Ctrl+Space — extend the current selection to span every row
 *  of every column it touches. No-op if there's no active range. */
export function selectEntireColumns(api: FUniver) {
  const range = activeRange(api);
  const sheet = activeSheet(api);
  if (!range || !sheet) return;
  const startColumn = range.getColumn();
  const endColumn = startColumn + range.getWidth() - 1;
  const lastRow = maxRows(sheet) - 1;
  const box = facadeRangeBox(sheet, {
    startRow: 0,
    endRow: lastRow,
    startColumn,
    endColumn,
  });
  if (box) activateRange(box);
}

/** Excel's Shift+Space — extend the current selection to span every
 *  column of every row it touches. */
export function selectEntireRows(api: FUniver) {
  const range = activeRange(api);
  const sheet = activeSheet(api);
  if (!range || !sheet) return;
  const startRow = range.getRow();
  const endRow = startRow + range.getHeight() - 1;
  const lastCol = maxColumns(sheet) - 1;
  const box = facadeRangeBox(sheet, {
    startRow,
    endRow,
    startColumn: 0,
    endColumn: lastCol,
  });
  if (box) activateRange(box);
}

/** Excel's F2 — enter edit mode on the active cell. Dispatches the
 *  Univer command that the canvas listens for. */
export function enterCellEditMode(api: FUniver) {
  void api.executeCommand('sheet.operation.start-edit', {});
}

/** Insert/Delete cells with one of the four Excel directions. The
 *  shift-cells variants leave neighbouring rows/cols untouched and
 *  push only the target range — Excel's Ctrl+Shift+= submenu. */
export type CellsOpDirection = 'shift-right' | 'shift-down' | 'entire-row' | 'entire-column';

export async function insertCellsAt(api: FUniver, dir: CellsOpDirection): Promise<void> {
  const range = activeRange(api);
  const wb = api.getActiveWorkbook();
  const sheet = activeSheet(api);
  if (!range || !wb || !sheet) return;
  if (dir === 'entire-row') {
    insertRowAbove(api);
    return;
  }
  if (dir === 'entire-column') {
    insertColumnLeft(api);
    return;
  }
  const cmd =
    dir === 'shift-down'
      ? 'sheet.command.insert-range-move-down'
      : 'sheet.command.insert-range-move-right';
  await api.executeCommand(cmd, {
    unitId: wb.getId(),
    subUnitId: facadeSheetId(sheet),
    range: rangeBox(range),
  });
}

export async function deleteCellsAt(api: FUniver, dir: CellsOpDirection): Promise<void> {
  const range = activeRange(api);
  const wb = api.getActiveWorkbook();
  const sheet = activeSheet(api);
  if (!range || !wb || !sheet) return;
  if (dir === 'entire-row') {
    deleteSelectedRow(api);
    return;
  }
  if (dir === 'entire-column') {
    deleteSelectedColumn(api);
    return;
  }
  const cmd =
    dir === 'shift-down'
      ? 'sheet.command.delete-range-move-up'
      : 'sheet.command.delete-range-move-left';
  await api.executeCommand(cmd, {
    unitId: wb.getId(),
    subUnitId: facadeSheetId(sheet),
    range: rangeBox(range),
  });
}

function rangeBox(range: FRange) {
  return {
    startRow: range.getRow(),
    startColumn: range.getColumn(),
    endRow: range.getRow() + range.getHeight() - 1,
    endColumn: range.getColumn() + range.getWidth() - 1,
  };
}

/** Switch to the previous visible sheet — Ctrl+PageUp in Excel. Skips
 *  hidden sheets (matches Excel's behavior). No-op if already on the
 *  first visible sheet. */
export function switchToPreviousSheet(api: FUniver) {
  switchSheetByDelta(api, -1);
}

/** Switch to the next visible sheet — Ctrl+PageDown in Excel. */
export function switchToNextSheet(api: FUniver) {
  switchSheetByDelta(api, +1);
}

function switchSheetByDelta(api: FUniver, delta: -1 | 1): void {
  const wb = api.getActiveWorkbook();
  const active = wb?.getActiveSheet();
  if (!wb || !active) return;
  const visible = wb.getSheets().filter((s) => !isHidden(s));
  if (visible.length <= 1) return;
  const activeId = facadeSheetId(active);
  if (!activeId) return;
  const idx = visible.findIndex((s) => facadeSheetId(s) === activeId);
  if (idx < 0) return;
  const nextIdx = idx + delta;
  if (nextIdx < 0 || nextIdx >= visible.length) return;
  wb.setActiveSheet(visible[nextIdx]);
}

/** Ctrl+Home — jump to A1 on the active sheet. */
export function jumpToFirstCell(api: FUniver) {
  const ws = api.getActiveWorkbook()?.getActiveSheet();
  if (!ws) return;
  ws.getRange(0, 0).activate();
}

/** Ctrl+End — jump to the last used cell on the active sheet. Excel's
 *  "last cell" is the bottom-right of the used range, NOT the literal
 *  last allocated cell (which would be far out at the workbook growth
 *  cap). Falls back to A1 if the sheet has no data. */
export function jumpToLastCell(api: FUniver) {
  const wb = api.getActiveWorkbook();
  const ws = wb?.getActiveSheet();
  if (!wb || !ws) return;
  let lastRow = 0;
  let lastCol = 0;
  let found = false;
  try {
    const snap = saveWorkbook(wb);
    const sId = facadeSheetId(ws);
    const sheet = sId ? snap?.sheets?.[sId] : undefined;
    const cellData = sheet?.cellData ?? {};
    // cellData is typed `{[row: number]: {...}}` upstream so
    // Object.keys()'s string return needs an indexer cast to look up.
    const cd = cellData as Record<string, Record<string, unknown>>;
    for (const r of Object.keys(cd)) {
      const rNum = Number(r);
      if (rNum > lastRow) lastRow = rNum;
      for (const c of Object.keys(cd[r] ?? {})) {
        const cNum = Number(c);
        if (cNum > lastCol) lastCol = cNum;
        found = true;
      }
    }
  } catch {
    /* fall through to A1 */
  }
  if (!found) {
    ws.getRange(0, 0).activate();
    return;
  }
  ws.getRange(lastRow, lastCol).activate();
}

/** Ctrl+; — insert today's date (yyyy-mm-dd) into the active cell. */
export function insertTodayDate(api: FUniver) {
  const ws = api.getActiveWorkbook()?.getActiveSheet();
  const range = ws?.getActiveRange();
  if (!range) return;
  const today = new Date();
  const pad = (n: number) => n.toString().padStart(2, '0');
  const v = `${today.getFullYear()}-${pad(today.getMonth() + 1)}-${pad(today.getDate())}`;
  range.setValue({ v });
}

/** Ctrl+Shift+: — insert current time (hh:mm:ss) into the active cell. */
export function insertCurrentTime(api: FUniver) {
  const ws = api.getActiveWorkbook()?.getActiveSheet();
  const range = ws?.getActiveRange();
  if (!range) return;
  const now = new Date();
  const pad = (n: number) => n.toString().padStart(2, '0');
  const v = `${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`;
  range.setValue({ v });
}

/**
 * Ctrl+' / Ctrl+Shift+' — copy from the cell directly above into the active
 * cell. The Ctrl+' variant copies the formula (preserving the `=…`), the
 * Ctrl+Shift+' variant copies the computed value (no formula). No-op on row 1.
 */
export function copyFromAbove(api: FUniver, mode: 'formula' | 'value') {
  const sheet = activeSheet(api);
  const range = activeRange(api);
  if (!sheet || !range) return;
  const row = range.getRow();
  if (row <= 0) return;
  const col = range.getColumn();
  const above = sheet.getRange(row - 1, col).getCellData();
  if (!above) return;
  const target = sheet.getRange(row, col);
  if (mode === 'formula' && typeof above.f === 'string' && above.f.length > 0) {
    target.setValue({ f: above.f });
  } else {
    // Value mode (or formula mode with no formula above) — strip the formula
    // and write the evaluated value so the new cell holds the literal.
    target.setValue({ v: above.v ?? null });
  }
}

/**
 * Hide whatever row(s) the active selection touches. Excel's Ctrl+9
 * doesn't require the user to select full rows first — a single-cell
 * selection still hides the row containing that cell. Univer's command
 * filters its `selections` to `rangeType === ROW` only, so we explicitly
 * pass a ROW-typed range derived from the cell selection.
 */
export function hideSelectedRows(api: FUniver) {
  const wb = api.getActiveWorkbook();
  const sheet = activeSheet(api);
  const range = activeRange(api);
  if (!wb || !sheet || !range) return;
  const startRow = range.getRow();
  const endRow = startRow + range.getHeight() - 1;
  const maxCol = (sheet as unknown as { getMaxColumns?: () => number }).getMaxColumns?.() ?? 1;
  api.executeCommand('sheet.command.set-rows-hidden', {
    unitId: wb.getId(),
    subUnitId: sheet.getSheetId(),
    // rangeType: 1 === RANGE_TYPE.ROW
    ranges: [
      { startRow, endRow, startColumn: 0, endColumn: Math.max(0, maxCol - 1), rangeType: 1 },
    ],
  });
}

export function unhideSelectedRows(api: FUniver) {
  const wb = api.getActiveWorkbook();
  const sheet = activeSheet(api);
  const range = activeRange(api);
  if (!wb || !sheet || !range) return;
  const startRow = range.getRow();
  const endRow = startRow + range.getHeight() - 1;
  const maxCol = (sheet as unknown as { getMaxColumns?: () => number }).getMaxColumns?.() ?? 1;
  api.executeCommand('sheet.command.set-specific-rows-visible', {
    unitId: wb.getId(),
    subUnitId: sheet.getSheetId(),
    ranges: [
      { startRow, endRow, startColumn: 0, endColumn: Math.max(0, maxCol - 1), rangeType: 1 },
    ],
  });
}

/**
 * Reveal every hidden row on the active sheet — recovery action when a filter
 * (sheet-level or table-level) has left rows stuck hidden after the filter UI
 * was dismissed. Univer's set-specific-rows-visible command is a no-op for
 * rows that are already visible, so passing the full row span is safe.
 */
export function showAllRows(api: FUniver) {
  const wb = api.getActiveWorkbook();
  const sheet = activeSheet(api);
  if (!wb || !sheet) return;
  const rowCount = (sheet as unknown as { getMaxRows: () => number }).getMaxRows();
  if (!rowCount) return;
  api.executeCommand('sheet.command.set-specific-rows-visible', {
    unitId: wb.getId(),
    subUnitId: sheet.getSheetId(),
    ranges: [
      // rangeType: 1 === RANGE_TYPE.ROW.
      { startRow: 0, endRow: rowCount - 1, startColumn: 0, endColumn: 0, rangeType: 1 },
    ],
  });
}

/** Mirror of hideSelectedRows for columns — Ctrl+0 path. */
export function hideSelectedColumns(api: FUniver) {
  const wb = api.getActiveWorkbook();
  const sheet = activeSheet(api);
  const range = activeRange(api);
  if (!wb || !sheet || !range) return;
  const startColumn = range.getColumn();
  const endColumn = startColumn + range.getWidth() - 1;
  const maxRow = (sheet as unknown as { getMaxRows?: () => number }).getMaxRows?.() ?? 1;
  api.executeCommand('sheet.command.set-col-hidden', {
    unitId: wb.getId(),
    subUnitId: sheet.getSheetId(),
    // rangeType: 2 === RANGE_TYPE.COLUMN
    ranges: [
      { startRow: 0, endRow: Math.max(0, maxRow - 1), startColumn, endColumn, rangeType: 2 },
    ],
  });
}

export function unhideSelectedColumns(api: FUniver) {
  const wb = api.getActiveWorkbook();
  const sheet = activeSheet(api);
  const range = activeRange(api);
  if (!wb || !sheet || !range) return;
  const startColumn = range.getColumn();
  const endColumn = startColumn + range.getWidth() - 1;
  const maxRow = (sheet as unknown as { getMaxRows?: () => number }).getMaxRows?.() ?? 1;
  // Univer's "show specific cols" command id is `set-col-visible-on-cols`
  // (asymmetric with the row variant `set-specific-rows-visible`). Using
  // the wrong id silently no-ops the unhide. Verified against
  // vendor/univer/packages/sheets/src/commands/commands/set-col-visible.command.ts:49.
  api.executeCommand('sheet.command.set-col-visible-on-cols', {
    unitId: wb.getId(),
    subUnitId: sheet.getSheetId(),
    ranges: [
      { startRow: 0, endRow: Math.max(0, maxRow - 1), startColumn, endColumn, rangeType: 2 },
    ],
  });
}

export function insertImage(api: FUniver) {
  api.executeCommand('sheet.command.insert-float-image');
}

/* ── Formulas — function inserts ────────────────────────────────────────── */

/**
 * Insert a function template at the active cell:
 *   - If selection is multi-cell, write `=FN(<selection>)` in the cell after
 *     (Excel AutoSum semantics).
 *   - If single-cell, write `=FN()` so the user can complete arguments.
 *
 * This is the same path as `applyAutoFunction` but generalized to any name.
 */
export function insertFunction(api: FUniver, name: string) {
  const sheet = activeSheet(api);
  const range = activeRange(api);
  if (!sheet || !range) return;

  const isMulti = range.getWidth() * range.getHeight() > 1;
  if (isMulti) {
    const targetRow = range.getRow() + range.getHeight();
    const targetCol = range.getColumn();
    const target = sheet.getRange(targetRow, targetCol);
    target.setValue({ f: `=${name}(${range.getA1Notation()})` });
    target.activate();
  } else {
    const cell = sheet.getRange(range.getRow(), range.getColumn());
    cell.setValue({ f: `=${name}()` });
  }
}

/* ── Data — Remove Duplicates / Text to Columns ─────────────────────────── */

/**
 * Remove duplicate rows in the active selection. Two rows are duplicates if
 * all their values (left→right) compare equal. The first occurrence is kept.
 * Cleared rows are blanked in place (Excel typically deletes; doing so via
 * the sheet ops is a bigger commit-undo dance and beyond MVP scope).
 */
export function removeDuplicates(api: FUniver) {
  const sheet = activeSheet(api);
  const range = activeRange(api);
  if (!sheet || !range) return;

  const startRow = range.getRow();
  const startCol = range.getColumn();
  const height = range.getHeight();
  const width = range.getWidth();
  const values = range.getValues();

  const seen = new Set<string>();
  for (let r = 0; r < height; r++) {
    const row = values[r] ?? [];
    const key = row.map((v: unknown) => stringify(v)).join('|');
    if (seen.has(key)) {
      // Wipe this row's values inside the selection.
      for (let c = 0; c < width; c++) {
        sheet.getRange(startRow + r, startCol + c).setValue({ v: null });
      }
    } else {
      seen.add(key);
    }
  }
}

function stringify(v: unknown): string {
  if (v === null || v === undefined) return '';
  if (typeof v === 'object' && 'v' in (v as Record<string, unknown>)) {
    return String((v as { v: unknown }).v ?? '');
  }
  return String(v);
}

export function splitTextToColumns(api: FUniver) {
  api.executeCommand('sheet.command.split-text-to-columns');
}

/* ── Review — comments ──────────────────────────────────────────────────── */

export function toggleCommentPanel(api: FUniver) {
  api.executeCommand('sheet.operation.toggle-comment-panel');
}

export function showCommentModal(api: FUniver) {
  api.executeCommand('sheet.operation.show-comment-modal');
}

/* ── Auto-fit ───────────────────────────────────────────────────────────── */

export function autoFitColumns(api: FUniver) {
  const range = activeRange(api);
  const sheet = activeSheet(api);
  if (!range || !sheet) return;
  // `setColumnAutoWidth` is on the sheets-ui facade extension.
  const sheetWithAutoWidth = sheet as unknown as {
    setColumnAutoWidth?: (col: number, n: number) => unknown;
  };
  sheetWithAutoWidth.setColumnAutoWidth?.(range.getColumn(), range.getWidth());
}

/** Soft cap on per-row autofit work — Univer's `autoFitRow` is O(cells
 *  in row) and we loop one call per row. A whole-column selection
 *  expands to the sheet's row count (often 8k+ even for blank sheets,
 *  millions after growth), which blocks the main thread for seconds.
 *  Excel's behavior on whole-column "Auto-fit Row Height" only touches
 *  rows that actually contain content; capping here matches that
 *  expectation closely enough without traversing every cell to detect
 *  emptiness. */
const AUTO_FIT_ROW_CAP = 500;

export function autoFitRows(api: FUniver) {
  const range = activeRange(api);
  const sheet = activeSheet(api);
  if (!range || !sheet) return;
  const start = range.getRow();
  const requested = range.getHeight();
  const count = Math.min(requested, AUTO_FIT_ROW_CAP);
  if (count < requested) {
    console.info(
      `[autofit-rows] capped ${requested} → ${AUTO_FIT_ROW_CAP} rows to keep the UI responsive`,
    );
  }
  for (let r = 0; r < count; r++) {
    sheet.autoFitRow(start + r);
  }
}

/* ── Formulas tab — quick auto-functions ─────────────────────────────────── */

type AutoFn = 'SUM' | 'AVERAGE' | 'COUNT' | 'MIN' | 'MAX';

/**
 * Mimics Excel's AutoSum behavior: if the selection is multi-cell, write the
 * formula in the cell immediately below or right of the selection. Otherwise
 * write the formula in the active cell with no range (user will fill it in).
 */
export function applyAutoFunction(api: FUniver, fn: AutoFn) {
  const sheet = activeSheet(api);
  const range = activeRange(api);
  if (!sheet || !range) return;

  const isMulti = range.getWidth() * range.getHeight() > 1;
  if (isMulti) {
    const targetRow = range.getRow() + range.getHeight();
    const targetCol = range.getColumn();
    const formula = `=${fn}(${range.getA1Notation()})`;
    const target = sheet.getRange(targetRow, targetCol);
    target.setValue({ f: formula });
    target.activate();
  } else {
    const cell = sheet.getRange(range.getRow(), range.getColumn());
    cell.setValue({ f: `=${fn}()` });
  }
}

/* ── Data tab — sort / filter ────────────────────────────────────────────── */

export function sortRange(api: FUniver, ascending: boolean) {
  const range = activeRange(api);
  if (!range) return;
  // sort() comes from sheets-sort facade. Cast for typing — the facade
  // augments FRange at runtime via FUniver.extend().
  const withSort = range as unknown as {
    sort?: (spec: { column: number; ascending: boolean }) => unknown;
  };
  withSort.sort?.({ column: range.getColumn(), ascending });
}

/**
 * Open Univer's Custom Sort panel for the active range — supports multi-key
 * sort (sort by A, then B, then C). The panel is shipped by sheets-sort-ui;
 * we dispatch the command rather than building our own dialog.
 */
export function openCustomSort(api: FUniver) {
  api.executeCommand('sheet.command.sort-range-custom');
}

/* ── Wave B — Insert / Data / View ──────────────────────────────────────── */

export function insertHyperlink(api: FUniver) {
  // Univer's hyper-link UI registers an "insert hyper-link" popup; the
  // command id below opens it with the active range pre-filled.
  api.executeCommand('sheet.operation.insert-hyper-link');
}

export function insertComment(api: FUniver) {
  // Opens the comment-on-cell modal — Univer figures out the cell from
  // the active selection.
  api.executeCommand('sheet.operation.show-comment-modal');
}

export async function insertTable(api: FUniver) {
  const range = activeRange(api);
  const sheet = activeSheet(api);
  const wb = api.getActiveWorkbook();
  if (!range || !sheet || !wb) return;
  // Table plugin is lazy-loaded; await before dispatch so the command
  // handler is registered. Without this, fast clicks on a fresh page
  // silently no-op (queued or dropped).
  await ensurePluginByName('table');
  api.executeCommand('sheet.command.add-table', {
    unitId: wb.getId(),
    subUnitId: sheet.getSheetId(),
    range: {
      startRow: range.getRow(),
      startColumn: range.getColumn(),
      endRow: range.getRow() + range.getHeight() - 1,
      endColumn: range.getColumn() + range.getWidth() - 1,
    },
  });
}

/**
 * Univer ships six built-in table themes: indigo, teal, green, purple,
 * pink, red. Surfaced as the Format-as-Table dropdown swatches.
 */
export const TABLE_THEMES = [
  { id: 'table-default-0', label: 'Indigo', swatch: '#6280F9' },
  { id: 'table-default-1', label: 'Teal', swatch: '#16BDCA' },
  { id: 'table-default-2', label: 'Green', swatch: '#31C48D' },
  { id: 'table-default-3', label: 'Purple', swatch: '#AC94FA' },
  { id: 'table-default-4', label: 'Pink', swatch: '#F17EBB' },
  { id: 'table-default-5', label: 'Red', swatch: '#F98080' },
] as const;

export type TableThemeId = (typeof TABLE_THEMES)[number]['id'];

/** Guard against double-fire while a previous formatAsTable is in flight
 *  (plugin lazy-load + addTable round-trip). Without this, two fast clicks
 *  on the "Format as Table" dropdown create two overlapping tables. */
let formatAsTableInFlight = false;

/**
 * Convert the active selection — or the contiguous data block around a
 * single-cell selection — into a styled Univer table, then auto-fit its
 * column widths. Mirrors Excel's "Format as Table".
 */
export async function formatAsTable(api: FUniver, themeId?: TableThemeId) {
  if (formatAsTableInFlight) return;
  formatAsTableInFlight = true;
  try {
    const wb = api.getActiveWorkbook();
    const sheet = activeSheet(api);
    const range = activeRange(api);
    if (!wb || !sheet || !range) return;

    const bounds = detectTableBounds(api, range);
    if (!bounds) return;

    // Table plugin is lazy-loaded — without this await, `addTable` is
    // undefined on the facade until the plugin finishes registering,
    // and a fast click is silently dropped. The in-flight guard above
    // covers the case where the plugin loads BETWEEN two clicks (the
    // second click finds `addTable` defined and creates a second table).
    await ensurePluginByName('table');

    // Second-fire guard: if a table already overlaps `bounds` on this
    // sheet, the user almost certainly clicked twice. Excel surfaces
    // an explicit error here; we just no-op (less intrusive). Without
    // this, a sequential rapid click — where the first click's table
    // has finished registering by the time the second click runs —
    // creates a duplicate. The `formatAsTableInFlight` boolean only
    // catches truly synchronous double-fires.
    // TODO(B1-followup): tables-plugin-specific API (`getTableList`,
    // `addTable`); not on the standard FWorkbook/FWorksheet facade.
    // Add typed wrappers to univer-facade.ts when we tackle plugin
    // surfaces in a follow-up pass.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const existing = (wb as any).getTableList?.() ?? [];
    const overlaps = existing.some(
      (t: { subUnitId?: string; range?: typeof bounds }) =>
        t.subUnitId === sheet.getSheetId() &&
        t.range &&
        !(
          t.range.endRow < bounds.startRow ||
          t.range.startRow > bounds.endRow ||
          t.range.endColumn < bounds.startColumn ||
          t.range.startColumn > bounds.endColumn
        ),
    );
    if (overlaps) {
      console.info('[format-as-table] skipped — a table already covers this range');
      return;
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const fws = sheet as any;
    if (typeof fws.addTable !== 'function') return;

    const id = `table-${Date.now().toString(36)}`;
    const name = `Table_${Date.now().toString(36)}`;
    const options = themeId ? { tableStyleId: themeId } : undefined;
    const result = fws.addTable(name, bounds, id, options);

    const after = () => {
      // Auto-fit the table's columns to their content — but only for small /
      // medium tables. setColumnAutoWidth measures (font-shapes) every cell in
      // the range, which is O(rows×cols): on a 21k-row table that's ~10s of pure
      // text measurement. Excel doesn't auto-fit on "Format as Table" anyway, so
      // for large ranges keep the imported widths instead of freezing the UI.
      const AUTO_WIDTH_CELL_CAP = 20_000;
      const cells =
        (bounds.endRow - bounds.startRow + 1) * (bounds.endColumn - bounds.startColumn + 1);
      if (cells > AUTO_WIDTH_CELL_CAP) return;
      const sheetAny = sheet as unknown as {
        setColumnAutoWidth?: (col: number, n: number) => unknown;
      };
      sheetAny.setColumnAutoWidth?.(bounds.startColumn, bounds.endColumn - bounds.startColumn + 1);
    };
    if (result && typeof (result as Promise<boolean>).then === 'function') {
      await (result as Promise<boolean>).then(after).catch(() => {});
    } else {
      after();
    }
  } finally {
    formatAsTableInFlight = false;
  }
}

/**
 * Expand a single-cell selection to the contiguous block of populated cells
 * around it (Excel's "current region"). Multi-cell selections are trusted.
 */
function detectTableBounds(
  api: FUniver,
  range: FRange,
): { startRow: number; startColumn: number; endRow: number; endColumn: number } | null {
  const isMulti = range.getWidth() * range.getHeight() > 1;
  if (isMulti) {
    return {
      startRow: range.getRow(),
      startColumn: range.getColumn(),
      endRow: range.getRow() + range.getHeight() - 1,
      endColumn: range.getColumn() + range.getWidth() - 1,
    };
  }

  const sheet = api.getActiveWorkbook()?.getActiveSheet();
  if (!sheet) return null;

  const startRow0 = range.getRow();
  const startCol0 = range.getColumn();

  const hasValue = (r: number, c: number): boolean => {
    try {
      const cd = sheet.getRange(r, c).getCellData();
      const v = cd?.v;
      return v !== undefined && v !== null && v !== '';
    } catch {
      return false;
    }
  };

  // Empty starting cell → return a 1×1 range so the user gets an empty table.
  if (!hasValue(startRow0, startCol0)) {
    return {
      startRow: startRow0,
      startColumn: startCol0,
      endRow: startRow0,
      endColumn: startCol0,
    };
  }

  let top = startRow0;
  while (top > 0 && hasValue(top - 1, startCol0)) top--;
  let bottom = startRow0;
  const maxRows = sheet.getMaxRows();
  while (bottom < maxRows - 1 && hasValue(bottom + 1, startCol0)) bottom++;
  let left = startCol0;
  while (left > 0 && hasValue(top, left - 1)) left--;
  let right = startCol0;
  const maxCols = sheet.getMaxColumns();
  while (right < maxCols - 1 && hasValue(top, right + 1)) right++;

  return { startRow: top, startColumn: left, endRow: bottom, endColumn: right };
}

export function openConditionalFormatting(api: FUniver) {
  // The panel command requires a `value` indicating the sub-mode.
  // 'viewRule' opens the all-rules manager view (Excel's default entry).
  api.executeCommand('sheet.operation.open.conditional.formatting.panel', {
    value: 'viewRule',
  });
}

export function openDataValidation(api: FUniver) {
  api.executeCommand('data-validation.operation.open-validation-panel');
}

/* ── View tab ───────────────────────────────────────────────────────────── */

// Univer's built-in set-first-row-frozen / set-first-column-frozen commands
// each clobber the orthogonal axis to zero, so they can't be used to freeze
// both a row and a column at the same time. The facade's setFrozenRows /
// setFrozenColumns preserve the other axis, which is what users expect from
// Excel (Freeze top row over a sheet that already has Freeze first column
// leaves both frozen).
//
// eslint-disable-next-line @typescript-eslint/no-explicit-any
type FreezeCapableSheet = {
  setFrozenRows: (n: number) => unknown;
  setFrozenColumns: (n: number) => unknown;
};

export function freezeFirstRow(api: FUniver) {
  const sheet = activeSheet(api) as unknown as FreezeCapableSheet | null;
  sheet?.setFrozenRows(1);
}

export function freezeFirstColumn(api: FUniver) {
  const sheet = activeSheet(api) as unknown as FreezeCapableSheet | null;
  sheet?.setFrozenColumns(1);
}

export function freezeAtSelection(api: FUniver) {
  api.executeCommand('sheet.command.set-selection-frozen');
}

export function unfreezePanes(api: FUniver) {
  api.executeCommand('sheet.command.cancel-frozen');
}

export function toggleGridlines(api: FUniver) {
  const wb = api.getActiveWorkbook();
  const sheet = wb?.getActiveSheet();
  if (!wb || !sheet) return;
  const sId = sheet.getSheetId();
  // Read the CURRENT state and flip it. (Was hardcoded `current=true`, so it
  // always sent showGridlines:0 — gridlines could be hidden but never shown.)
  // BooleanNumber: 1 = show (default when unset), 0 = hide.
  const snap = saveWorkbook(wb);
  const cfg = snap?.sheets?.[sId] as { showGridlines?: number } | undefined;
  const shown = cfg?.showGridlines !== 0;
  api.executeCommand('sheet.command.toggle-gridlines', {
    unitId: wb.getId(),
    subUnitId: sId,
    showGridlines: shown ? 0 : 1,
  });
}

export function setZoom(api: FUniver, ratio: number) {
  const wb = api.getActiveWorkbook();
  const sheet = wb?.getActiveSheet();
  if (!wb || !sheet) return;
  api.executeCommand('sheet.command.set-zoom-ratio', {
    unitId: wb.getId(),
    subUnitId: sheet.getSheetId(),
    zoomRatio: Math.max(0.1, Math.min(4, ratio)),
  });
}

export async function toggleFilter(api: FUniver): Promise<void> {
  // sheets-filter is lazy-loaded; await before touching the facade.
  await ensurePluginByName('filter');
  const wb = api.getActiveWorkbook();
  const sheet = activeSheet(api);
  const range = activeRange(api);
  if (!wb || !sheet || !range) return;

  // Dispatch the underlying commands directly rather than through the
  // facade's createFilter / remove(). The facade uses syncExecuteCommand
  // which returns false silently when the plugin's command handlers
  // haven't finished registering after a fresh dynamic import. The
  // async executeCommand path waits for the dispatch and surfaces the
  // result. Command ids verified against
  // vendor/univer/packages/sheets-filter/src/commands/commands/sheets-filter.command.ts.
  const sheetWithFilter = sheet as unknown as { getFilter?: () => unknown };
  const existing = sheetWithFilter.getFilter?.();
  if (existing) {
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
}

/**
 * Excel's Ctrl+Alt+L — re-apply the active filter. After editing cells
 * inside a filtered range, rows that no longer match the criteria stay
 * visible until the filter is re-evaluated. This re-runs that pass.
 *
 * Command id verified against
 * vendor/univer/packages/sheets-filter/src/commands/commands/sheets-filter.command.ts:282.
 * The handler resolves its target sheet from the instance service when
 * params are omitted. No-ops cleanly when the sheet has no filter.
 */
export async function reapplyFilter(api: FUniver): Promise<void> {
  await ensurePluginByName('filter');
  const wb = api.getActiveWorkbook();
  const sheet = activeSheet(api);
  if (!wb || !sheet) return;
  await api.executeCommand('sheet.command.re-calc-filter', {
    unitId: wb.getId(),
    subUnitId: sheet.getSheetId(),
  });
}

/* ── Formulas tab — force recalc ────────────────────────────────────────── */

/**
 * Excel's F9 — re-run the whole dependency graph even for cells whose
 * inputs didn't change. Useful after editing a volatile UDF or when a
 * referenced external source has refreshed and the engine missed it.
 *
 * The mutation id is the one engine-formula listens on internally
 * (`vendor/univer/packages/engine-formula/src/commands/mutations/
 * set-formula-calculation.mutation.ts:83`); passing `forceCalculation`
 * skips the dependency short-circuit.
 */
export function forceRecalculate(api: FUniver) {
  api.executeCommand('formula.mutation.set-formula-calculation-start', {
    forceCalculation: true,
  });
}
