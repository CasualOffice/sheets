import { useEffect } from 'react';
import { useUniverAPI } from '../use-univer';
import { MAX_COLUMNS, MAX_ROWS } from '../snapshot';

/**
 * Dynamic workbook growth. When the user navigates / scrolls / selects near
 * the bottom or right edge of the sheet, append rows/columns in chunks up
 * to a hard cap (MAX_ROWS / MAX_COLUMNS).
 *
 * Why this approach
 *   Declaring a huge initial size (8192 × 1024) materializes a lot of
 *   row/column metadata up front and slows boot. Growing on demand keeps
 *   the initial workbook small while still feeling unbounded to the user.
 */

const EDGE_BUFFER_ROWS = 32; // grow when within this many of the current max
const EDGE_BUFFER_COLS = 8;
const GROW_ROWS_CHUNK = 256;
const GROW_COLS_CHUNK = 32;

const SELECTION_OP_ID = 'sheet.operation.set-selections';

export function useWorkbookGrowth() {
  const api = useUniverAPI();

  useEffect(() => {
    if (!api) return;

    const tryGrow = () => {
      const wb = api.getActiveWorkbook();
      const sheet = wb?.getActiveSheet();
      if (!sheet) return;

      const selection = sheet.getActiveRange();
      if (!selection) return;

      const lastRow = selection.getRow() + selection.getHeight() - 1;
      const lastCol = selection.getColumn() + selection.getWidth() - 1;

      const maxRows = sheet.getMaxRows();
      const maxCols = sheet.getMaxColumns();

      // Use setRowCount / setColumnCount (Facade) — direct count bump, no
      // per-row undo overhead the way insertRowsAfter would incur.
      if (lastRow >= maxRows - EDGE_BUFFER_ROWS && maxRows < MAX_ROWS) {
        const next = Math.min(maxRows + GROW_ROWS_CHUNK, MAX_ROWS);
        sheet.setRowCount(next);
      }

      if (lastCol >= maxCols - EDGE_BUFFER_COLS && maxCols < MAX_COLUMNS) {
        const next = Math.min(maxCols + GROW_COLS_CHUNK, MAX_COLUMNS);
        sheet.setColumnCount(next);
      }
    };

    // Initial check (in case a snapshot loads with a large initial selection).
    tryGrow();

    const disposable = api.addEvent(api.Event.CommandExecuted, (e) => {
      if ((e as { id?: string }).id === SELECTION_OP_ID) tryGrow();
    });
    return () => disposable.dispose();
  }, [api]);
}
