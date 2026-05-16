import { useEffect, useState } from 'react';
import { useUniverAPI } from '../use-univer';

/**
 * Live state of the active cell + current selection, derived from Univer events.
 * Drives ribbon toggle states (Bold/Italic/Underline/Alignment) and the
 * status bar selection stats.
 *
 * Subscribes to `SelectionChanged` and `SheetValueChanged` — together these
 * cover the "user moved the cursor" and "value/style changed in current cell"
 * cases that should re-render the chrome.
 */

export type HAlign = 'left' | 'center' | 'right' | 'unset';

export type ActiveCellState = {
  ready: boolean;
  isBold: boolean;
  isItalic: boolean;
  isUnderline: boolean;
  align: HAlign;
  numberFormat: string;
  /** Selection-level numeric aggregates (excludes single-cell selection). */
  stats: { count: number; sum: number; avg: number | null } | null;
};

const EMPTY: ActiveCellState = {
  ready: false,
  isBold: false,
  isItalic: false,
  isUnderline: false,
  align: 'unset',
  numberFormat: '',
  stats: null,
};

export function useActiveCellState(): ActiveCellState {
  const api = useUniverAPI();
  const [state, setState] = useState<ActiveCellState>(EMPTY);

  useEffect(() => {
    if (!api) return;

    const compute = (): ActiveCellState => {
      const wb = api.getActiveWorkbook();
      if (!wb) return EMPTY;
      const sheet = wb.getActiveSheet();
      if (!sheet) return EMPTY;
      const selection = wb.getActiveSheet().getActiveRange();
      if (!selection) return EMPTY;

      // Active cell = top-left of current selection.
      const row = selection.getRow();
      const col = selection.getColumn();
      const cell = sheet.getRange(row, col);
      const cellData = cell.getCellData();
      const style =
        typeof cellData?.s === 'string'
          ? wb.getWorkbook().getStyles().get(cellData.s) ?? null
          : (cellData?.s ?? null);

      // Selection stats (multi-cell only).
      let stats: ActiveCellState['stats'] = null;
      const cellsX = selection.getWidth();
      const cellsY = selection.getHeight();
      if (cellsX * cellsY > 1) {
        let count = 0;
        let sum = 0;
        const values = selection.getValues();
        for (const row of values) {
          for (const v of row) {
            const inner =
              typeof v === 'object' && v !== null && 'v' in (v as Record<string, unknown>)
                ? (v as { v: unknown }).v
                : v;
            const n = typeof inner === 'number' ? inner : Number(inner);
            if (!Number.isNaN(n) && inner !== null && inner !== '' && inner !== undefined) {
              count += 1;
              sum += n;
            }
          }
        }
        stats = { count, sum, avg: count > 0 ? sum / count : null };
      }

      return {
        ready: true,
        isBold: style?.bl === 1,
        isItalic: style?.it === 1,
        isUnderline: !!style?.ul && (style.ul.s ?? 0) === 1,
        align:
          style?.ht === 1 ? 'left' : style?.ht === 2 ? 'center' : style?.ht === 3 ? 'right' : 'unset',
        numberFormat: style?.n?.pattern ?? '',
        stats,
      };
    };

    setState(compute());

    /*
     * Two events on their own don't cover everything we want to react to:
     *   - `SelectionChanged` fires only for canvas-driven changes; programmatic
     *     FRange.activate() / setActiveRange go through `SetSelectionsOperation`
     *     but bypass `selectionChanged$`.
     *   - `SheetValueChanged` fires only when cell values change — style-only
     *     mutations (numfmt, bold, alignment) don't fire it.
     *
     * The reliable signal is `CommandExecuted` filtered to:
     *   - any `sheet.mutation.*` (covers value + style + numfmt changes)
     *   - the selection operation id (covers programmatic selection changes)
     *
     * Important for both user UX and any future scripted / AI command surface.
     */
    const SET_SELECTIONS_OP_ID = 'sheet.operation.set-selections';
    const shouldRecompute = (id: string | undefined) =>
      !!id && (id.startsWith('sheet.mutation.') || id === SET_SELECTIONS_OP_ID);

    const disposable = api.addEvent(api.Event.CommandExecuted, (e) => {
      if (shouldRecompute((e as { id?: string }).id)) {
        setState(compute());
      }
    });
    return () => disposable.dispose();
  }, [api]);

  return state;
}
