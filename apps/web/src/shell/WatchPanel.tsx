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

import { useEffect, useMemo, useState } from 'react';
import type { FUniver } from '@univerjs/core/facade';
import { useUniverAPI } from '../use-univer';
import { useUI } from '../use-ui';
import { Icon } from './Icon';
import { useWatches } from './watch-context';
import { cellA1, cellsInRect, type Watch } from './watch-model';

/**
 * Watch Window (Excel's Formulas → Watch Window). Pins cells so their value +
 * formula stay visible while you scroll or switch sheets. "Add watch" pins the
 * current selection; the list re-reads live values on every edit. Builds on the
 * shared `.side-panel` rail chrome.
 */

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function sheetById(api: FUniver | null, sheetId: string): any | null {
  const wb = api?.getActiveWorkbook();
  if (!wb) return null;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  return (wb.getSheets() as any[]).find((s) => s.getSheetId?.() === sheetId) ?? null;
}

type Resolved = Watch & { value: string; formula: string };

export function WatchPanel() {
  const api = useUniverAPI();
  const ui = useUI();
  const { watches, add, remove, clear } = useWatches();
  // Bumped on edits so the displayed values re-read.
  const [tick, setTick] = useState(0);

  useEffect(() => {
    if (!api) return;
    const disp = api.addEvent(api.Event.CommandExecuted, (e) => {
      const id = (e as { id?: string }).id ?? '';
      // Re-read on any cell mutation (direct edits AND formula recalc results
      // flowing back from the worker) plus workbook swaps. Re-reading a handful
      // of cells is cheap, so a broad trigger keeps watched values honest.
      if (id.startsWith('sheet.mutation.') || id === 'doc.command-replace-snapshot') {
        setTick((t) => t + 1);
      }
    });
    return () => disp.dispose();
  }, [api]);

  const rows = useMemo<Resolved[]>(() => {
    return watches.map((w) => {
      const ws = sheetById(api, w.sheetId);
      let value = '';
      let formula = '';
      try {
        const cell = ws?.getRange?.(w.row, w.col);
        const v = cell?.getValue?.();
        value = v == null ? '' : String(v);
        formula = cell?.getFormula?.() ?? '';
      } catch {
        /* sheet gone / not ready — leave blank */
      }
      return { ...w, value, formula };
    });
    // `tick` forces a live re-read on edits.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [watches, api, tick]);

  const addSelection = () => {
    if (!api) return;
    const wb = api.getActiveWorkbook();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws = wb?.getActiveSheet() as any;
    const range = ws?.getActiveRange?.();
    const sheetId = ws?.getSheetId?.();
    if (!range || sheetId == null) return;
    const rect = {
      startRow: range.getRow(),
      startColumn: range.getColumn(),
      endRow: range.getRow() + (range.getHeight?.() ?? 1) - 1,
      endColumn: range.getColumn() + (range.getWidth?.() ?? 1) - 1,
    };
    const sheetName = ws.getSheetName?.() ?? '';
    add(cellsInRect(rect).map(({ row, col }) => ({ sheetId, sheetName, row, col })));
  };

  // Double-click a watch → jump to that cell (Excel's Watch Window behaviour),
  // switching sheets if it lives on another one.
  const navigate = (w: Watch) => {
    const wb = api?.getActiveWorkbook();
    const sheet = sheetById(api, w.sheetId);
    if (!wb || !sheet) return;
    try {
      wb.setActiveSheet?.(sheet);
      sheet.getRange?.(w.row, w.col)?.activate?.();
    } catch {
      /* sheet removed since the watch was added */
    }
  };

  const empty = watches.length === 0;

  return (
    <aside className="side-panel watch-panel" data-testid="watch-panel">
      <header className="side-panel__header">
        <Icon name="visibility" size="sm" />
        <h2 className="side-panel__title">Watch Window</h2>
        {!empty && <span className="side-panel__count">{watches.length}</span>}
        <button
          type="button"
          className="side-panel__close"
          aria-label="Close Watch Window"
          onClick={ui.toggleWatchPanel}
        >
          <Icon name="close" size="sm" />
        </button>
      </header>

      <div className="side-panel__body watch-panel__body">
        <div className="watch-panel__actions">
          <button
            type="button"
            className="btn-secondary"
            data-testid="watch-add"
            disabled={!api}
            onClick={addSelection}
          >
            <Icon name="add" size="sm" /> Add watch
          </button>
          {!empty && (
            <button
              type="button"
              className="btn-secondary"
              data-testid="watch-clear"
              onClick={clear}
            >
              Clear all
            </button>
          )}
        </div>

        {empty ? (
          <div className="side-panel__empty" data-testid="watch-empty">
            <Icon name="visibility" size="lg" className="side-panel__empty-icon" />
            <div className="side-panel__empty-title">No watched cells</div>
            <div className="side-panel__empty-body">
              Select cells and choose <strong>Add watch</strong> to keep an eye on their values and
              formulas as you work.
            </div>
          </div>
        ) : (
          <table className="watch-panel__table" data-testid="watch-table">
            <thead>
              <tr>
                <th>Sheet</th>
                <th>Cell</th>
                <th>Value</th>
                <th>Formula</th>
                <th aria-label="Remove" />
              </tr>
            </thead>
            <tbody>
              {rows.map((w) => (
                <tr
                  key={w.id}
                  data-testid={`watch-row-${w.id}`}
                  className="watch-panel__row"
                  title="Double-click to go to this cell"
                  onDoubleClick={() => navigate(w)}
                >
                  <td title={w.sheetName}>{w.sheetName}</td>
                  <td>{cellA1(w.row, w.col)}</td>
                  <td data-testid={`watch-value-${cellA1(w.row, w.col)}`}>{w.value}</td>
                  <td className="watch-panel__formula" title={w.formula}>
                    {w.formula}
                  </td>
                  <td>
                    <button
                      type="button"
                      className="watch-panel__remove"
                      aria-label={`Remove watch on ${w.sheetName}!${cellA1(w.row, w.col)}`}
                      data-testid={`watch-remove-${cellA1(w.row, w.col)}`}
                      onClick={() => remove(w.id)}
                    >
                      <Icon name="close" size="sm" />
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>
    </aside>
  );
}
