import type { FUniver } from '@univerjs/core/facade';
import { ISheetClipboardService } from '@univerjs/sheets-ui';

/**
 * Excel-paste merge preservation.
 *
 * Univer's HTML clipboard parser reads `colspan` / `rowspan` from
 * pasted `<td>` cells into the cell matrix, but the default
 * `onPasteCells` hook only uses those numbers to *skip* covered
 * cells — it never emits `AddWorksheetMergeMutation`. So a paste of
 * a merged Excel range lands as N separate cells in our grid.
 *
 * This function registers a higher-priority clipboard hook that runs
 * AFTER the default hook (Univer sorts hooks ascending by priority,
 * so a larger number runs later — see the sort at
 * `@univerjs/sheets-ui` ES bundle line ~6228) and walks the
 * destination matrix for any cell with `rowSpan > 1` / `colSpan > 1`,
 * emitting `sheet.mutation.add-worksheet-merge` for each.
 *
 * We register via the FUniver facade's `_injector` rather than as a
 * proper Plugin class because the project's tsconfig doesn't enable
 * `experimentalDecorators` and Univer's Plugin DI uses parameter
 * decorators. The injector reach matches the existing pattern in
 * `apps/web/src/shell/file-actions.ts:179` for `IMessageService`.
 *
 * Returns a disposer the caller can call on teardown.
 */
export function registerPasteMergeHook(api: FUniver): (() => void) | null {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const injector = (api as any)._injector as
    | { get: (token: unknown) => unknown }
    | undefined;
  if (!injector) return null;
  let clipboard: unknown;
  try {
    clipboard = injector.get(ISheetClipboardService);
  } catch {
    return null;
  }
  const svc = clipboard as
    | {
        addClipboardHook: (hook: ClipboardHook) => { dispose: () => void };
      }
    | undefined;
  if (!svc?.addClipboardHook) return null;

  const onPasteCells: ClipboardHook['onPasteCells'] = (_pasteFrom, pasteTo, data) => {
    const ranges: MergeRange[] = [];
    data.forValue?.((r: number, c: number, cell: SpanCell | null) => {
      const rs = cell?.rowSpan ?? 1;
      const cs = cell?.colSpan ?? 1;
      if (rs <= 1 && cs <= 1) return;
      const rows = pasteTo.range.rows;
      const cols = pasteTo.range.cols;
      const startRow = rows[r];
      const startCol = cols[c];
      if (startRow === undefined || startCol === undefined) return;
      // Discrete ranges allow non-contiguous indices (e.g. filtered
      // destination). Fall back to the contiguous case if the
      // expected index isn't present.
      const endRow = rows[r + rs - 1] ?? startRow + rs - 1;
      const endCol = cols[c + cs - 1] ?? startCol + cs - 1;
      ranges.push({
        startRow,
        endRow,
        startColumn: startCol,
        endColumn: endCol,
      });
    });
    if (ranges.length === 0) return { undos: [], redos: [] };
    const params = {
      unitId: pasteTo.unitId,
      subUnitId: pasteTo.subUnitId,
      ranges,
    };
    return {
      redos: [{ id: 'sheet.mutation.add-worksheet-merge', params }],
      undos: [{ id: 'sheet.mutation.remove-worksheet-merge', params }],
    };
  };

  // Expose the bare callback in dev so e2e specs can exercise its
  // mutation-generation logic without having to drive a real
  // ClipboardEvent (synthetic events don't reach Univer's listener in
  // headless Chromium). Tree-shaken in production via the
  // `import.meta.env.DEV` gate.
  if (import.meta.env.DEV) {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (globalThis as any).__pasteMergeHook__ = onPasteCells;
  }

  // Companion hook for column widths. Univer's HTML parser already
  // reads `<col width="…">` into `colProperties`, but the default
  // paste hook only honors them on the "special paste / column
  // widths" path — not on a regular Ctrl+V. Without this, pasting
  // a styled Excel table loses its column sizing.
  //
  // Row heights have the same problem upstream but `onPasteRows` is
  // called with no `rowProperties` argument (verified against the
  // sheets-ui ES bundle's paste loop at line ~4418), so there's no
  // way to recover them from a hook. That gap needs a pnpm patch or
  // an upstream PR.
  const onPasteColumns: ClipboardHook['onPasteColumns'] = (
    pasteTo,
    colProperties,
    _payload,
  ) => {
    if (!Array.isArray(colProperties) || colProperties.length === 0) {
      return { undos: [], redos: [] };
    }
    const cols = pasteTo.range.cols;
    const widthsByCol: Record<number, number> = {};
    let any = false;
    for (let i = 0; i < colProperties.length; i++) {
      const destCol = cols[i];
      if (destCol === undefined) continue;
      const raw = colProperties[i]?.width;
      if (typeof raw !== 'string' && typeof raw !== 'number') continue;
      // Strip any trailing "px"/"pt" and round; Excel typically writes
      // a bare number but some sources include units.
      const n = Math.round(Number.parseFloat(String(raw)));
      if (!Number.isFinite(n) || n <= 0) continue;
      widthsByCol[destCol] = n;
      any = true;
    }
    if (!any) return { undos: [], redos: [] };
    const destColIndices = Object.keys(widthsByCol).map((k) => Number(k));
    const startColumn = Math.min(...destColIndices);
    const endColumn = Math.max(...destColIndices);
    const params = {
      unitId: pasteTo.unitId,
      subUnitId: pasteTo.subUnitId,
      ranges: [{ startRow: 0, endRow: 0, startColumn, endColumn }],
      colWidth: widthsByCol,
    };
    return {
      redos: [{ id: 'sheet.mutation.set-worksheet-col-width', params }],
      // Undo restores via the same mutation with the inverse map; we
      // could read the current widths here, but the destination is
      // freshly pasted-into so the safest "undo" is to no-op the
      // change. A subsequent Ctrl+Z still rewinds the cell content
      // mutations the default hook emitted.
      undos: [],
    };
  };

  // Row heights — symmetrical to col widths. The upstream sheets-ui
  // dispatcher used to call `onPasteRows` with no rowProperties
  // argument (a real bug — `rowProperties` IS parsed from the HTML
  // but never forwarded). We carry a pnpm patch
  // (`patches/@univerjs__sheets-ui@0.22.1.patch`) that adds the
  // second/third args. With the patch applied, this hook receives
  // the same shape as onPasteColumns.
  const onPasteRows: ClipboardHook['onPasteRows'] = (
    pasteTo,
    rowProperties,
    _payload,
  ) => {
    if (!Array.isArray(rowProperties) || rowProperties.length === 0) {
      return { undos: [], redos: [] };
    }
    const rows = pasteTo.range.rows;
    const heightsByRow: Record<number, number> = {};
    let any = false;
    for (let i = 0; i < rowProperties.length; i++) {
      const destRow = rows[i];
      if (destRow === undefined) continue;
      const raw = rowProperties[i]?.height;
      if (typeof raw !== 'string' && typeof raw !== 'number') continue;
      const n = Math.round(Number.parseFloat(String(raw)));
      if (!Number.isFinite(n) || n <= 0) continue;
      heightsByRow[destRow] = n;
      any = true;
    }
    if (!any) return { undos: [], redos: [] };
    const destRowIndices = Object.keys(heightsByRow).map((k) => Number(k));
    const startRow = Math.min(...destRowIndices);
    const endRow = Math.max(...destRowIndices);
    const params = {
      unitId: pasteTo.unitId,
      subUnitId: pasteTo.subUnitId,
      ranges: [{ startRow, endRow, startColumn: 0, endColumn: 0 }],
      rowHeight: heightsByRow,
    };
    return {
      redos: [{ id: 'sheet.mutation.set-worksheet-row-height', params }],
      undos: [],
    };
  };

  if (import.meta.env.DEV) {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (globalThis as any).__pasteColWidthHook__ = onPasteColumns;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (globalThis as any).__pasteRowHeightHook__ = onPasteRows;
  }

  const mergeDisposable = svc.addClipboardHook({
    id: 'casual-sheets-paste-merges',
    priority: 1000,
    onPasteCells,
  });
  const colWidthDisposable = svc.addClipboardHook({
    id: 'casual-sheets-paste-col-widths',
    priority: 1000,
    onPasteColumns,
  });
  const rowHeightDisposable = svc.addClipboardHook({
    id: 'casual-sheets-paste-row-heights',
    priority: 1000,
    onPasteRows,
  });
  return () => {
    mergeDisposable.dispose();
    colWidthDisposable.dispose();
    rowHeightDisposable.dispose();
    if (import.meta.env.DEV) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      delete (globalThis as any).__pasteMergeHook__;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      delete (globalThis as any).__pasteColWidthHook__;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      delete (globalThis as any).__pasteRowHeightHook__;
    }
  };
}

// Minimal local types for the bits of Univer's clipboard hook we touch.
// Keeps the file independent of @univerjs/sheets-ui type re-exports.
type MergeRange = {
  startRow: number;
  endRow: number;
  startColumn: number;
  endColumn: number;
};
type SpanCell = { rowSpan?: number; colSpan?: number };
type DiscreteRange = { rows: number[]; cols: number[] };
type PasteTarget = { range: DiscreteRange; unitId: string; subUnitId: string };
type DataMatrix = {
  forValue?: (
    cb: (r: number, c: number, cell: SpanCell | null) => void,
  ) => unknown;
};
type MutationInfo = { id: string; params: unknown };
type ColProperty = Record<string, string | undefined>;
type ClipboardHook = {
  id: string;
  priority?: number;
  onPasteCells?: (
    pasteFrom: unknown,
    pasteTo: PasteTarget,
    data: DataMatrix,
    payload?: unknown,
  ) => { undos: MutationInfo[]; redos: MutationInfo[] };
  onPasteColumns?: (
    pasteTo: PasteTarget,
    colProperties: ColProperty[],
    payload: unknown,
  ) => { undos: MutationInfo[]; redos: MutationInfo[] };
  // Upstream sheets-ui originally called this with only `pasteTo`; our
  // patch adds the extra args matching onPasteColumns.
  onPasteRows?: (
    pasteTo: PasteTarget,
    rowProperties: ColProperty[],
    payload: unknown,
  ) => { undos: MutationInfo[]; redos: MutationInfo[] };
};
