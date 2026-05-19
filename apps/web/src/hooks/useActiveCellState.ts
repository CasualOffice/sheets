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
export type VAlign = 'top' | 'middle' | 'bottom' | 'unset';

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
  isStrike: boolean;
  isWrapped: boolean;
  fontFamily: string;
  fontSize: number;
  fontColor: string;
  fillColor: string;
  align: HAlign;
  vAlign: VAlign;
  numberFormat: string;
  /** True when the current selection exactly matches a merged range. */
  isMerged: boolean;
  /** True when the current selection spans more than one cell. */
  isMultiCell: boolean;
  /** Rows × columns of the primary selection. Shown next to the Name
   *  Box when `isMultiCell`. Excel shows "3R x 2C" transiently during
   *  drag; we surface it permanently so the info doesn't blink away. */
  selRows: number;
  selCols: number;
  /** True while the Format Painter is armed (one-shot or infinite). */
  isFormatPainterActive: boolean;
  /** Selection-level numeric aggregates. Covers every selected range
   *  (Ctrl+click adds to the list), matching Excel's status bar
   *  semantics. Null when the selection is a single empty cell. */
  stats: {
    count: number;
    sum: number;
    avg: number | null;
    min: number | null;
    max: number | null;
    /** Total cell count across all selected ranges — Excel's "Count"
     *  (non-empty cells) vs "Numerical Count" (count). We expose the
     *  numerical count via `count` and the cell count via `cellCount`
     *  so the status bar can show both. */
    cellCount: number;
  } | null;
};

const EMPTY: ActiveCellState = {
  ready: false,
  a1: '',
  formula: '',
  displayValue: '',
  isBold: false,
  isItalic: false,
  isUnderline: false,
  isStrike: false,
  isWrapped: false,
  fontFamily: '',
  fontSize: 11,
  fontColor: '',
  fillColor: '',
  align: 'unset',
  vAlign: 'unset',
  numberFormat: '',
  isMerged: false,
  isMultiCell: false,
  selRows: 0,
  selCols: 0,
  isFormatPainterActive: false,
  stats: null,
};

const SET_SELECTIONS_OP_ID = 'sheet.operation.set-selections';
const FORMAT_PAINTER_OP_ID = 'sheet.operation.set-format-painter';
const shouldRecompute = (id: string | undefined) =>
  !!id &&
  (id.startsWith('sheet.mutation.') ||
    id === SET_SELECTIONS_OP_ID ||
    id === FORMAT_PAINTER_OP_ID);

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

      // Format Painter armed state — Univer exposes it on a service we can
      // read via the injector. Status 0 = off, 1 = one-shot armed, 2 = infinite.
      let isFormatPainterActive = false;
      try {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const injector = (api as any)._univer?.__getInjector?.();
        const svc = injector?.get?.({ name: 'IFormatPainterService' });
        const status = svc?.getStatus?.();
        isFormatPainterActive = status !== undefined && status !== 0;
      } catch {
        // service may not exist; leave as false
      }

      // Selection stats — aggregate across EVERY selected range
      // (Ctrl+click adds disjoint regions; Excel's status bar treats
      // them as one bag). Cap the materialization at 100k cells —
      // beyond that, `getValues()` allocates a 2D array of millions
      // of entries which freezes the UI for seconds when a user hits
      // Cmd+A or clicks the select-all corner on a big workbook. The
      // status-bar stat isn't worth a 2 s freeze; we surface it as
      // null and the UI hides the row.
      const SELECTION_STATS_CAP = 100_000;
      let stats: ActiveCellState['stats'] = null;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const sel = (sheet as any).getSelection?.();
      const ranges: { getValues: () => unknown[][]; getWidth: () => number; getHeight: () => number }[] =
        sel?.getActiveRangeList?.() ?? [selection];
      let totalCells = 0;
      for (const r of ranges) totalCells += r.getWidth() * r.getHeight();
      if (totalCells > 0 && totalCells <= SELECTION_STATS_CAP) {
        let cellCount = 0;
        let count = 0;
        let sum = 0;
        let min = Number.POSITIVE_INFINITY;
        let max = Number.NEGATIVE_INFINITY;
        for (const r of ranges) {
          const values = r.getValues();
          for (const row of values) {
            for (const v of row) {
              const inner =
                typeof v === 'object' && v !== null && 'v' in (v as Record<string, unknown>)
                  ? (v as { v: unknown }).v
                  : v;
              if (inner === null || inner === undefined || inner === '') continue;
              cellCount += 1;
              const n = typeof inner === 'number' ? inner : Number(inner);
              if (Number.isFinite(n)) {
                count += 1;
                sum += n;
                if (n < min) min = n;
                if (n > max) max = n;
              }
            }
          }
        }
        // For a single-cell selection containing a non-numeric value
        // (or a single number) Excel still shows Count=1; we only
        // suppress stats when nothing was selected at all.
        if (cellCount > 0) {
          stats = {
            count,
            sum,
            avg: count > 0 ? sum / count : null,
            min: count > 0 ? min : null,
            max: count > 0 ? max : null,
            cellCount,
          };
        }
      }

      return {
        ready: true,
        a1: cell.getA1Notation(),
        formula,
        displayValue,
        isBold: style?.bl === 1,
        isItalic: style?.it === 1,
        isUnderline: !!style?.ul && (style.ul.s ?? 0) === 1,
        isStrike: !!style?.st && (style.st.s ?? 0) === 1,
        isWrapped: cell.getWrap(),
        fontFamily: style?.ff ?? '',
        fontSize: style?.fs ?? 11,
        fontColor: (style?.cl && typeof style.cl === 'object' && 'rgb' in style.cl ? style.cl.rgb : '') ?? '',
        fillColor: (style?.bg && typeof style.bg === 'object' && 'rgb' in style.bg ? style.bg.rgb : '') ?? '',
        align:
          style?.ht === 1 ? 'left' : style?.ht === 2 ? 'center' : style?.ht === 3 ? 'right' : 'unset',
        vAlign:
          style?.vt === 1 ? 'top' : style?.vt === 2 ? 'middle' : style?.vt === 3 ? 'bottom' : 'unset',
        numberFormat: style?.n?.pattern ?? '',
        isMerged,
        isMultiCell,
        selRows: selH,
        selCols: selW,
        isFormatPainterActive,
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
