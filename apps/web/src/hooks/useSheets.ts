import { useEffect, useState } from 'react';
import { useUniverAPI } from '../use-univer';

/**
 * Live list of sheets in the active workbook, plus the active sheet id.
 * Reacts to: SheetCreated, SheetDeleted, SheetNameChanged, ActiveSheetChanged.
 */

export type SheetSummary = { id: string; name: string; hidden: boolean };

export type SheetsState = {
  ready: boolean;
  sheets: SheetSummary[];
  activeSheetId: string | null;
};

const EMPTY: SheetsState = { ready: false, sheets: [], activeSheetId: null };

export function useSheets(): SheetsState {
  const api = useUniverAPI();
  const [state, setState] = useState<SheetsState>(EMPTY);

  useEffect(() => {
    if (!api) return;

    const compute = (): SheetsState => {
      const wb = api.getActiveWorkbook();
      if (!wb) return EMPTY;
      const sheets = wb.getSheets().map((s) => ({
        id: s.getSheetId(),
        name: s.getSheetName(),
        // `isSheetHidden` is on the sheets facade — hidden xlsx sheets
        // round-trip through here. The tab strip uses this to skip
        // them from the visible row and offer a separate "Unhide…"
        // menu, matching Excel behavior.
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        hidden: (s as any).isSheetHidden?.() === true,
      }));
      const active = wb.getActiveSheet();
      return {
        ready: true,
        sheets,
        activeSheetId: active?.getSheetId() ?? null,
      };
    };

    setState(compute());

    const refresh = () => setState(compute());
    const disposables = [
      api.addEvent(api.Event.SheetCreated, refresh),
      api.addEvent(api.Event.SheetDeleted, refresh),
      api.addEvent(api.Event.SheetNameChanged, refresh),
      api.addEvent(api.Event.ActiveSheetChanged, refresh),
      api.addEvent(api.Event.SheetMoved, refresh),
      // File → Open / drag-drop swap the entire workbook unit. None of
      // the SheetCreated / SheetMoved events fire for those — the sheets
      // come in as part of the new unit's initial snapshot. Subscribe to
      // unit creation so the tab strip refreshes when the user opens a
      // multi-sheet xlsx. Deferred a tick because Univer emits
      // unitAdded$ BEFORE setCurrentUnitForType — calling compute()
      // synchronously here would still see the old active workbook.
      api.onUniverSheetCreated(() => queueMicrotask(refresh)),
      // Visibility (hide/show) doesn't have a dedicated facade event —
      // the underlying mutation `sheet.mutation.set-worksheet-hidden`
      // fires through CommandExecuted. Subscribe so the tab strip
      // immediately removes/restores tabs when a sheet is hidden or
      // unhidden from the context menu.
      api.addEvent(api.Event.CommandExecuted, (e) => {
        const id = (e as { id?: string }).id;
        if (id === 'sheet.mutation.set-worksheet-hidden') refresh();
      }),
    ];
    return () => {
      for (const d of disposables) d.dispose();
    };
  }, [api]);

  return state;
}
