import { useEffect, useState } from 'react';
import { useUniverAPI } from '../use-univer';

/**
 * Live state of the active cell + current selection, derived from Univer events.
 * Drives ribbon toggle states, the formula bar, and the status bar stats.
 *
 * Two events alone don't cover every case we want to react to:
 *   - `SelectionChanged` fires only for canvas-driven changes; programmatic
 *     FRange.activate() / setActiveRange go through `SetSelectionsOperation`
 *     but bypass `selectionChanged$`.
 *   - `SheetValueChanged` fires only when cell values change — style-only
 *     mutations (numfmt, bold, alignment) don't fire it.
 *
 * The reliable signal is `CommandExecuted`, filtered to:
 *   - any `sheet.mutation.*` (covers value + style + numfmt changes)
 *   - the selection operation id (covers programmatic selection changes)
 *
 * Important for both user UX and any future scripted / AI command surface.
 */

export type HAlign = 'left' | 'center' | 'right' | 'unset';

export type ActiveCellState = {
  ready: boolean;
  /** A1 reference of the active cell (top-left of the selection). */
  a1: string;
  /** Formula text without the leading `=` stripped — empty string if no formula. */
  formula: string;
  /** Display string: formula if present, otherwise the raw value. */
  displayValue: string;
  isBold: boolean;
  isItalic: boolean;
  isUnderline: boolean;
  align: HAlign;
  numberFormat: string;
  /** True when the current selection exactly matches a merged range. */
  isMerged: boolean;
  /** True when the current selection spans more than one cell. */
  isMultiCell: boolean;
  /** Selection-level numeric aggregates (excludes single-cell selection). */
  stats: { count: number; sum: number; avg: number | null } | null;
};

const EMPTY: ActiveCellState = {
  ready: false,
  a1: '',
  formula: '',
  displayValue: '',
  isBold: false,
  isItalic: false,
  isUnderline: false,
  align: 'unset',
  numberFormat: '',
  isMerged: false,
  isMultiCell: false,
  stats: null,
};

const SET_SELECTIONS_OP_ID = 'sheet.operation.set-selections';
const shouldRecompute = (id: string | undefined) =>
  !!id && (id.startsWith('sheet.mutation.') || id === SET_SELECTIONS_OP_ID);

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
      const selection = sheet.getActiveRange();
      if (!selection) return EMPTY;

      // Active cell = top-left of current selection.
      const row = selection.getRow();
      const col = selection.getColumn();
      const cell = sheet.getRange(row, col);
      const cellData = cell.getCellData();
      const style =
        typeof cellData?.s === 'string'
          ? (wb.getWorkbook().getStyles().get(cellData.s) ?? null)
          : (cellData?.s ?? null);

      const formula = cellData?.f ?? '';
      const rawValue = cellData?.v;
      const displayValue =
        formula !== ''
          ? formula
          : rawValue === null || rawValue === undefined
            ? ''
            : String(rawValue);

      // Detect merge: does the current selection exactly match one of the
      // worksheet's merged ranges?
      const selRow = selection.getRow();
      const selCol = selection.getColumn();
      const selW = selection.getWidth();
      const selH = selection.getHeight();
      const merges = sheet.getMergedRanges();
      const isMerged = merges.some((m) => {
        const r = m.getRange();
        return (
          r.startRow === selRow &&
          r.startColumn === selCol &&
          r.endRow === selRow + selH - 1 &&
          r.endColumn === selCol + selW - 1
        );
      });
      const isMultiCell = selW * selH > 1;

      // Selection stats (multi-cell only).
      let stats: ActiveCellState['stats'] = null;
      const cellsX = selW;
      const cellsY = selH;
      if (cellsX * cellsY > 1) {
        let count = 0;
        let sum = 0;
        const values = selection.getValues();
        for (const r of values) {
          for (const v of r) {
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
        a1: cell.getA1Notation(),
        formula,
        displayValue,
        isBold: style?.bl === 1,
        isItalic: style?.it === 1,
        isUnderline: !!style?.ul && (style.ul.s ?? 0) === 1,
        align:
          style?.ht === 1 ? 'left' : style?.ht === 2 ? 'center' : style?.ht === 3 ? 'right' : 'unset',
        numberFormat: style?.n?.pattern ?? '',
        isMerged,
        isMultiCell,
        stats,
      };
    };

    setState(compute());

    const disposable = api.addEvent(api.Event.CommandExecuted, (e) => {
      if (shouldRecompute((e as { id?: string }).id)) {
        setState(compute());
      }
    });
    return () => disposable.dispose();
  }, [api]);

  return state;
}
