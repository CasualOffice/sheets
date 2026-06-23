import type { FUniver } from '@univerjs/core/facade';
import { ISheetClipboardService } from '@univerjs/sheets-ui';
import { MAX_COLUMNS, MAX_ROWS } from '../snapshot';

/**
 * Paste-to-fit grid growth (Phase 1, T1.3).
 *
 * The grid materializes a modest number of rows/columns and grows on demand as
 * the user navigates near an edge (`useWorkbookGrowth`). But a paste places
 * content relative to the anchor cell, not the current selection — Univer's
 * clipboard service expands the destination to the clipboard's dimensions
 * (`_expandOrShrinkRowsCols`, which does NOT clamp to the worksheet bound). So
 * pasting a block taller/wider than the current extent lands rows/cols beyond
 * `rowCount`/`columnCount`: the cells exist in the sparse matrix but are
 * invisible because the grid only renders up to the declared extent.
 *
 * `onBeforePaste` receives the full (un-clamped) destination range and runs
 * before the paste mutations, so we grow the sheet to cover it first — capped at
 * MAX_ROWS / MAX_COLUMNS. The destination rows/cols are ascending (the expand
 * loop pushes increasing indices), so the last entry is the max — avoids a
 * `Math.max(...huge)` spread blowing the call stack on a large paste.
 *
 * Registered via the FUniver `_injector` (same pattern as `paste-merge-hook.ts`)
 * because the tsconfig doesn't enable `experimentalDecorators`. Never blocks a
 * paste: any failure returns `true` so the paste proceeds unchanged.
 */
type DiscretePasteTo = {
  unitId: string;
  subUnitId: string;
  range: { rows: number[]; cols: number[] };
};

interface PasteGrowthHook {
  id: string;
  priority?: number;
  onBeforePaste?: (pasteTo: DiscretePasteTo) => boolean;
}

/**
 * Grow the active sheet to cover a paste destination. Exported (and exposed as
 * `window.__pasteGrowthHook__` in dev) so e2e can exercise it directly — a real
 * Ctrl+V / clipboard paste doesn't reach Univer's listener in headless Chromium
 * (same constraint documented in `paste-merge-hook.ts`).
 */
export function growToFitPaste(api: FUniver, pasteTo: DiscretePasteTo): boolean {
  try {
    const rows = pasteTo?.range?.rows;
    const cols = pasteTo?.range?.cols;
    const ws = api.getActiveWorkbook()?.getActiveSheet();
    if (!ws) return true;

    if (Array.isArray(rows) && rows.length > 0) {
      const neededRows = rows[rows.length - 1] + 1; // ascending → last is max
      if (neededRows > ws.getMaxRows()) {
        ws.setRowCount(Math.min(neededRows, MAX_ROWS));
      }
    }
    if (Array.isArray(cols) && cols.length > 0) {
      const neededCols = cols[cols.length - 1] + 1;
      if (neededCols > ws.getMaxColumns()) {
        ws.setColumnCount(Math.min(neededCols, MAX_COLUMNS));
      }
    }
  } catch {
    // Growth is best-effort — never block a paste on it.
  }
  return true;
}

export function registerPasteGrowthHook(api: FUniver): (() => void) | null {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const injector = (api as any)._injector as { get: (token: unknown) => unknown } | undefined;
  if (!injector) return null;

  let clipboard: unknown;
  try {
    clipboard = injector.get(ISheetClipboardService);
  } catch {
    return null;
  }
  const svc = clipboard as
    | { addClipboardHook?: (hook: PasteGrowthHook) => { dispose: () => void } }
    | undefined;
  if (!svc?.addClipboardHook) return null;

  const onBeforePaste = (pasteTo: DiscretePasteTo) => growToFitPaste(api, pasteTo);

  // Expose the callback in dev so e2e can drive it without a real clipboard event.
  if (import.meta.env.DEV) {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (globalThis as any).__pasteGrowthHook__ = onBeforePaste;
  }

  const disposable = svc.addClipboardHook({
    id: 'casual-paste-growth',
    // Run early so the grid is grown before any hook that reads the extent.
    priority: 10000,
    onBeforePaste,
  });

  return () => disposable.dispose();
}
