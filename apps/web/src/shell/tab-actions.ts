import type { FUniver } from '@univerjs/core/facade';
import type { FRange } from '@univerjs/sheets/facade';

/** Imperative dispatchers used by Insert / Formulas / Data tabs. */

function activeRange(api: FUniver) {
  const wb = api.getActiveWorkbook();
  const sheet = wb?.getActiveSheet();
  return sheet?.getActiveRange() ?? null;
}

function activeSheet(api: FUniver) {
  return api.getActiveWorkbook()?.getActiveSheet() ?? null;
}

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

export function hideSelectedRows(api: FUniver) {
  api.executeCommand('sheet.command.set-rows-hidden');
}

export function unhideSelectedRows(api: FUniver) {
  api.executeCommand('sheet.command.set-selected-rows-visible');
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

export function hideSelectedColumns(api: FUniver) {
  api.executeCommand('sheet.command.set-col-hidden');
}

export function unhideSelectedColumns(api: FUniver) {
  api.executeCommand('sheet.command.set-selected-cols-visible');
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
    const key = row.map((v) => stringify(v)).join('|');
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

export function autoFitRows(api: FUniver) {
  const range = activeRange(api);
  const sheet = activeSheet(api);
  if (!range || !sheet) return;
  // FWorksheet.autoFitRow takes a single row index. Loop the selection.
  const start = range.getRow();
  for (let r = 0; r < range.getHeight(); r++) {
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

export function insertTable(api: FUniver) {
  const range = activeRange(api);
  const sheet = activeSheet(api);
  const wb = api.getActiveWorkbook();
  if (!range || !sheet || !wb) return;
  // `sheet.command.add-table` is the public command from sheets-table.
  // It needs unitId / subUnitId / range params.
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

/**
 * Convert the active selection — or the contiguous data block around a
 * single-cell selection — into a styled Univer table, then auto-fit its
 * column widths. Mirrors Excel's "Format as Table".
 */
export function formatAsTable(api: FUniver, themeId?: TableThemeId) {
  const wb = api.getActiveWorkbook();
  const sheet = activeSheet(api);
  const range = activeRange(api);
  if (!wb || !sheet || !range) return;

  const bounds = detectTableBounds(api, range);
  if (!bounds) return;

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const fws = sheet as any;
  if (typeof fws.addTable !== 'function') return;

  const id = `table-${Date.now().toString(36)}`;
  const name = `Table_${Date.now().toString(36)}`;
  const options = themeId ? { tableStyleId: themeId } : undefined;
  const result = fws.addTable(name, bounds, id, options);

  const after = () => {
    const sheetAny = sheet as unknown as {
      setColumnAutoWidth?: (col: number, n: number) => unknown;
    };
    sheetAny.setColumnAutoWidth?.(
      bounds.startColumn,
      bounds.endColumn - bounds.startColumn + 1,
    );
  };
  if (result && typeof (result as Promise<boolean>).then === 'function') {
    (result as Promise<boolean>).then(after).catch(() => {});
  } else {
    after();
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

export function toggleGridlines(api: FUniver, current: boolean) {
  const wb = api.getActiveWorkbook();
  const sheet = wb?.getActiveSheet();
  if (!wb || !sheet) return;
  // BooleanNumber: 0 = hide, 1 = show
  api.executeCommand('sheet.command.toggle-gridlines', {
    unitId: wb.getId(),
    subUnitId: sheet.getSheetId(),
    showGridlines: current ? 0 : 1,
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

export function toggleFilter(api: FUniver) {
  const sheet = activeSheet(api);
  const range = activeRange(api);
  if (!sheet || !range) return;

  // sheets-filter facade augmentations.
  const sheetWithFilter = sheet as unknown as { getFilter?: () => { remove: () => void } | null };
  const existing = sheetWithFilter.getFilter?.();
  if (existing) {
    existing.remove();
    return;
  }
  const rangeWithCreateFilter = range as unknown as { createFilter?: () => unknown };
  rangeWithCreateFilter.createFilter?.();
}
