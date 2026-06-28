/**
 * Copyright 2026 Casual Office
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

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

    // Facade events (SheetCreated, SheetDeleted, SheetNameChanged,
    // SheetMoved, ActiveSheetChanged) only fire when the corresponding
    // COMMAND runs. The collab bridge replays peer changes by invoking
    // the MUTATION directly via `executeCommand(mutation.id, ...,
    // { fromCollab: true })`, which means peers never see those facade
    // events for remote changes — Univer's internal state updates but
    // the tab strip stayed stale until something else forced a refresh.
    //
    // Belt-and-braces: also subscribe to CommandExecuted for every
    // mutation id that affects the sheet list. Catches both the
    // command path (own edits) AND the bridge's mutation-only replay
    // path (peer edits).
    const SHEET_LIST_MUTATIONS = new Set([
      'sheet.mutation.insert-sheet',
      'sheet.mutation.remove-sheet',
      'sheet.mutation.set-worksheet-name',
      'sheet.mutation.set-worksheet-order',
      'sheet.mutation.set-worksheet-hidden',
    ]);

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
      // Mutation-level fallback for cross-peer + visibility changes.
      // See the SHEET_LIST_MUTATIONS comment above.
      api.addEvent(api.Event.CommandExecuted, (e) => {
        const id = (e as { id?: string }).id;
        if (id && SHEET_LIST_MUTATIONS.has(id)) refresh();
      }),
    ];
    return () => {
      for (const d of disposables) d.dispose();
    };
  }, [api]);

  return state;
}
