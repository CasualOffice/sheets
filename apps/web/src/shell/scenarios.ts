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

/**
 * Scenario Manager — model + ref parsing (Excel's What-If → Scenario Manager).
 *
 * A scenario is a named snapshot of the values of a set of "changing cells".
 * Showing a scenario writes those values back into the cells; the sheet
 * recalculates naturally. Scenarios are kept per workbook in a session store
 * (lost on reload — persisting them into the workbook is a follow-up).
 *
 * The ref parsing + store are framework-free so they're unit-testable; the
 * dialog reads/writes current cell values through the facade.
 */

export interface ScenarioCell {
  /** A1 reference of a single changing cell, e.g. "B2". */
  ref: string;
  /** The cell's value in this scenario (string or number, as read/written). */
  value: string | number;
}

export interface Scenario {
  name: string;
  cells: ScenarioCell[];
}

/**
 * Parse a "changing cells" string ("B1, B2 C3") into a deduped list of
 * single-cell A1 refs (upper-cased). Ranges and malformed tokens are dropped.
 */
export function parseCellRefs(input: string): string[] {
  const seen = new Set<string>();
  const out: string[] = [];
  for (const tokenRaw of input.split(/[\s,]+/)) {
    const token = tokenRaw.trim().toUpperCase();
    if (token === '') continue;
    if (!/^[A-Z]+[0-9]+$/.test(token)) continue; // single cell only
    if (seen.has(token)) continue;
    seen.add(token);
    out.push(token);
  }
  return out;
}

// ── Per-workbook session store ───────────────────────────────────────────────
// Keyed by unitId so distinct workbooks don't share scenarios. Module-level so
// it survives the dialog opening/closing within a session.
const store = new Map<string, Scenario[]>();

export function getScenarios(unitId: string): Scenario[] {
  return store.get(unitId) ?? [];
}

export function setScenarios(unitId: string, scenarios: Scenario[]): void {
  store.set(unitId, scenarios);
}

/** Replace-or-append a scenario by name; returns the new list. */
export function upsertScenario(unitId: string, scenario: Scenario): Scenario[] {
  const list = getScenarios(unitId).filter((s) => s.name !== scenario.name);
  list.push(scenario);
  setScenarios(unitId, list);
  return list;
}

export function deleteScenario(unitId: string, name: string): Scenario[] {
  const list = getScenarios(unitId).filter((s) => s.name !== name);
  setScenarios(unitId, list);
  return list;
}

/** Test-only: clear all stored scenarios. */
export function __clearScenarios(): void {
  store.clear();
}
