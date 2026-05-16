import type { FUniver } from '@univerjs/core/facade';

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

export function freezeFirstRow(api: FUniver) {
  api.executeCommand('sheet.command.set-first-row-frozen');
}

export function freezeFirstColumn(api: FUniver) {
  api.executeCommand('sheet.command.set-first-column-frozen');
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
