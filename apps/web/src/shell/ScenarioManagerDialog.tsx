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

import { useState } from 'react';
import type { FUniver } from '@univerjs/core/facade';
import { Dialog } from './Dialog';
import {
  deleteScenario,
  getScenarios,
  parseCellRefs,
  upsertScenario,
  type Scenario,
} from './scenarios';

/**
 * Scenario Manager (Excel's What-If → Scenario Manager). A scenario is a named
 * snapshot of a set of "changing cells"; showing it writes those values back
 * and the sheet recalculates. Adding a scenario captures the cells' current
 * values, so the flow is: set the cells how you want → Add → name it → repeat,
 * then Show to flip between them. The model + store live in `scenarios.ts`.
 *
 * Session-scoped for now (kept per workbook in memory) — persisting scenarios
 * into the saved workbook is a follow-up.
 */

type Props = {
  api: FUniver;
  onClose: () => void;
};

export function ScenarioManagerDialog({ api, onClose }: Props) {
  const wb = api.getActiveWorkbook();
  const unitId: string = wb?.getId?.() ?? '';
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const ws = wb?.getActiveSheet() as any;

  const [list, setList] = useState<Scenario[]>(() => getScenarios(unitId));
  const [adding, setAdding] = useState<{ name: string; refs: string } | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [shown, setShown] = useState<string | null>(null);

  const create = () => {
    if (!adding || !ws) return;
    const name = adding.name.trim();
    if (!name) return setError('Name the scenario.');
    const refs = parseCellRefs(adding.refs);
    if (refs.length === 0) return setError('Enter one or more changing cells, e.g. B1, B2.');
    // Capture the cells' current values as this scenario.
    const cells = refs.map((ref) => {
      const v = ws.getRange(ref).getValue();
      return { ref, value: typeof v === 'number' ? v : (v ?? '') };
    });
    setList(upsertScenario(unitId, { name, cells }));
    setAdding(null);
    setError(null);
  };

  const show = (s: Scenario) => {
    if (!ws) return;
    for (const c of s.cells) {
      ws.getRange(c.ref).setValue({ v: c.value });
    }
    setShown(s.name);
  };

  const remove = (name: string) => {
    setList(deleteScenario(unitId, name));
    if (shown === name) setShown(null);
  };

  return (
    <Dialog
      title="Scenario Manager"
      onClose={onClose}
      data-testid="scenario-manager-dialog"
      footer={
        adding ? (
          <>
            <button
              type="button"
              className="btn-secondary"
              data-testid="scenario-add-cancel"
              onClick={() => {
                setAdding(null);
                setError(null);
              }}
            >
              Cancel
            </button>
            <button
              type="button"
              className="btn-primary"
              data-testid="scenario-add-save"
              onClick={create}
            >
              Add
            </button>
          </>
        ) : (
          <>
            <button
              type="button"
              className="btn-primary"
              data-testid="scenario-add"
              onClick={() => {
                const sel = ws?.getActiveRange?.()?.getA1Notation?.() ?? '';
                setAdding({ name: '', refs: sel });
              }}
            >
              Add…
            </button>
            <button
              type="button"
              className="btn-secondary"
              data-testid="scenario-close"
              onClick={onClose}
            >
              Close
            </button>
          </>
        )
      }
    >
      {adding ? (
        <div style={{ display: 'grid', gap: 10 }}>
          <label style={{ display: 'grid', gap: 4, fontSize: 13 }}>
            Scenario name
            <input
              autoFocus
              type="text"
              data-testid="scenario-name-input"
              value={adding.name}
              onChange={(e) => setAdding({ ...adding, name: e.target.value })}
              placeholder="Best case"
            />
          </label>
          <label style={{ display: 'grid', gap: 4, fontSize: 13 }}>
            Changing cells
            <input
              type="text"
              data-testid="scenario-cells-input"
              value={adding.refs}
              spellCheck={false}
              onChange={(e) => setAdding({ ...adding, refs: e.target.value })}
              placeholder="B1, B2"
            />
          </label>
          <div style={{ fontSize: 11, color: 'var(--cs-chrome-muted, #8a8886)' }}>
            Captures the current values of those cells. Set the cells how you want this scenario to
            look before adding it.
          </div>
          {error && (
            <div data-testid="scenario-error" style={{ fontSize: 12, color: '#b00020' }}>
              {error}
            </div>
          )}
        </div>
      ) : list.length === 0 ? (
        <div data-testid="scenario-empty" style={{ fontSize: 13 }}>
          No scenarios yet. Set your input cells, then click <strong>Add…</strong> to capture them
          as a named scenario.
        </div>
      ) : (
        <div style={{ display: 'grid', gap: 4 }}>
          {list.map((s) => (
            <div
              key={s.name}
              data-testid="scenario-row"
              style={{
                display: 'flex',
                alignItems: 'center',
                gap: 8,
                padding: '6px 8px',
                borderRadius: 6,
                background: shown === s.name ? 'var(--cs-chrome-active, #e6f3f7)' : 'transparent',
              }}
            >
              <span style={{ flex: 1, fontSize: 13 }}>
                {s.name}
                <span style={{ color: 'var(--cs-chrome-muted, #8a8886)', fontSize: 11 }}>
                  {' '}
                  ({s.cells.map((c) => c.ref).join(', ')})
                </span>
              </span>
              <button
                type="button"
                className="btn-secondary"
                data-testid="scenario-show"
                onClick={() => show(s)}
              >
                Show
              </button>
              <button
                type="button"
                className="btn-secondary"
                data-testid="scenario-delete"
                onClick={() => remove(s.name)}
              >
                Delete
              </button>
            </div>
          ))}
        </div>
      )}
    </Dialog>
  );
}
