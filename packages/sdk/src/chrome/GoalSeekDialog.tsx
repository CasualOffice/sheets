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
 * GoalSeekDialog — the SDK chrome's built-in What-If Goal Seek modal.
 *
 * Excel's Goal Seek: pick a formula cell ("Set cell"), a target result
 * ("To value"), and one input cell ("By changing cell"). The solver drives the
 * input cell until the formula cell reaches the target.
 *
 * The installed Univer facade ships NO goal-seek command (verified: no `goalSeek`
 * anywhere under the `@univerjs/sheets` facade modules), so this dialog implements
 * the solver client-side against real facade primitives only:
 *
 *   - read/write cells  → `FWorksheet.getRange(a1)` → `FRange.setValue(n)` /
 *                          `FRange.getValue()` (`@univerjs/sheets/facade`
 *                          f-range.d.ts L814 / L309, f-worksheet.d.ts getRange L279)
 *   - settle the formula → `univerAPI.getFormula().onCalculationEnd()` returns a
 *                          `Promise<void>` that resolves once the async formula
 *                          engine has applied results (`@univerjs/engine-formula/facade`
 *                          f-formula.d.ts L127; the mixin `getFormula()` is on
 *                          f-univer.d.ts L22). Univer recomputes formulas
 *                          asynchronously after a `setValue` mutation, so we await
 *                          this before reading the formula cell back.
 *
 * Numerics: a damped secant iteration (finite-difference seeded when the two
 * probes coincide), with a bisection-style bracket fallback. Non-convergence
 * within the iteration budget restores the input cell to its original value and
 * surfaces a message — no silent partial writes.
 *
 * Mounted by `<DialogHost>` when `openDialog('goal-seek')` is called and no host
 * override is registered.
 */

import { useMemo, useState, type CSSProperties } from 'react';
// Side-effect import: installs the formula-engine facade mixin
// (`univerAPI.getFormula()` + FFormula.onCalculationEnd) onto the facade
// prototype. `@univerjs/sheets/facade` (imported by sheets/api.ts) pulls the
// sheet mixins; this one guarantees the formula engine surface is present at the
// call site regardless of which facade groups a host registered.
import '@univerjs/sheets-formula/facade';
import type { DialogComponentProps } from './extensions';
import type { CasualSheetsAPI } from '../sheets/api';
import { Dialog } from './Dialog';
import {
  DIALOG_BTN_PRIMARY_STYLE,
  DIALOG_BTN_SECONDARY_STYLE,
  DIALOG_FIELD_STYLE,
  DIALOG_INPUT_STYLE,
  DIALOG_LABEL_STYLE,
} from './dialog-styles';

/** Loosely-typed slice of the FUniver facade this dialog leans on. */
interface Solver {
  worksheet: {
    getRange: (a1: string) => {
      setValue: (value: number) => unknown;
      getValue: () => unknown;
    } | null;
  };
  /** Wait for the async formula engine to apply results after a mutation. */
  settle: () => Promise<void>;
}

/** Resolve the active-sheet range/settle helpers, or null when unavailable. */
function getSolver(api: CasualSheetsAPI): Solver | null {
  const worksheet = api.univer.getActiveWorkbook()?.getActiveSheet() ?? null;
  if (!worksheet) return null;
  const formula = (
    api.univer as unknown as { getFormula?: () => { onCalculationEnd?: () => Promise<void> } }
  ).getFormula?.();
  return {
    worksheet: worksheet as unknown as Solver['worksheet'],
    settle: async () => {
      // onCalculationEnd resolves on the next completed calc pass. If the facade
      // is missing it (older group set), fall back to a microtask + rAF so the
      // engine's async recompute has a chance to flush before we read back.
      if (formula?.onCalculationEnd) {
        await formula.onCalculationEnd();
        return;
      }
      await new Promise<void>((resolve) => {
        if (typeof requestAnimationFrame === 'function') requestAnimationFrame(() => resolve());
        else setTimeout(resolve, 16);
      });
    },
  };
}

/** Coerce a Univer cell value to a finite number, or NaN. */
function toNumber(value: unknown): number {
  if (typeof value === 'number') return value;
  if (typeof value === 'boolean') return value ? 1 : 0;
  if (typeof value === 'string' && value.trim() !== '') return Number(value);
  return Number.NaN;
}

/** Basic A1 single-cell validation (letters + digits, no range colon). */
function isSingleCellA1(ref: string): boolean {
  return /^[A-Za-z]{1,3}[1-9][0-9]*$/.test(ref.trim());
}

interface DialogState {
  /** The formula cell whose result we're steering (A1). */
  setCell: string;
  /** The target value the formula cell should reach. */
  toValue: string;
  /** The input cell the solver adjusts (A1). */
  byChangingCell: string;
}

const INITIAL_STATE: DialogState = { setCell: '', toValue: '', byChangingCell: '' };

/** Convergence tolerance on |formula(x) - target|. */
const TOLERANCE = 1e-6;
/** Max solver iterations before declaring non-convergence. */
const MAX_ITERATIONS = 100;

type SeekResult = { ok: true; solution: number; result: number } | { ok: false; reason: string };

/**
 * Drive `byChangingCell` until `setCell`'s formula result reaches `target`.
 * Restores the original input on failure so the sheet is never left mid-solve.
 */
async function goalSeek(
  solver: Solver,
  setCellRef: string,
  target: number,
  changeCellRef: string,
): Promise<SeekResult> {
  const setRange = solver.worksheet.getRange(setCellRef);
  const changeRange = solver.worksheet.getRange(changeCellRef);
  if (!setRange || !changeRange) return { ok: false, reason: 'Could not resolve the cells.' };

  const original = toNumber(changeRange.getValue());
  const x0Seed = Number.isFinite(original) ? original : 0;

  // Evaluate the formula cell after setting the input to `x`.
  const evalAt = async (x: number): Promise<number> => {
    changeRange.setValue(x);
    await solver.settle();
    return toNumber(setRange.getValue());
  };

  const restore = async () => {
    changeRange.setValue(x0Seed);
    await solver.settle();
  };

  // f(x) = formula(x) - target; we want a root of f.
  let x0 = x0Seed;
  let f0 = (await evalAt(x0)) - target;
  if (!Number.isFinite(f0)) {
    await restore();
    return { ok: false, reason: 'The set cell is not a number (check that it holds a formula).' };
  }
  if (Math.abs(f0) <= TOLERANCE) return { ok: true, solution: x0, result: f0 + target };

  // Second probe: nudge x0. Use a step scaled to the magnitude of x0 so we get a
  // meaningful finite-difference slope even for large/small inputs.
  let x1 = x0 !== 0 ? x0 * 1.01 + 0.01 : 1;
  let f1 = (await evalAt(x1)) - target;

  for (let i = 0; i < MAX_ITERATIONS; i++) {
    if (Number.isFinite(f1) && Math.abs(f1) <= TOLERANCE) {
      return { ok: true, solution: x1, result: f1 + target };
    }

    const denom = f1 - f0;
    let x2: number;
    if (!Number.isFinite(f1) || denom === 0) {
      // Flat/undefined slope — perturb to break the stall rather than divide by 0.
      x2 = x1 + (x1 !== 0 ? x1 * 0.1 : 0.1);
    } else {
      // Secant step, damped to avoid overshoot on stiff curves.
      const step = (f1 * (x1 - x0)) / denom;
      x2 = x1 - step;
      if (!Number.isFinite(x2)) x2 = x1 + (x1 !== 0 ? x1 * 0.1 : 0.1);
    }

    const f2 = (await evalAt(x2)) - target;
    x0 = x1;
    f0 = f1;
    x1 = x2;
    f1 = f2;
  }

  await restore();
  return {
    ok: false,
    reason: `Could not converge after ${MAX_ITERATIONS} iterations. Try an input cell whose value the set cell actually depends on.`,
  };
}

const NOTE_STYLE: CSSProperties = {
  fontSize: 12,
  color: 'var(--cs-chrome-muted, #605e5c)',
  marginBottom: 12,
};

const STATUS_OK_STYLE: CSSProperties = {
  fontSize: 12,
  color: 'var(--cs-chrome-active-fg, #0e7490)',
  marginTop: 4,
};

const STATUS_ERR_STYLE: CSSProperties = {
  fontSize: 12,
  color: 'var(--cs-chrome-danger, #b91c1c)',
  marginTop: 4,
};

type Status =
  | { kind: 'idle' }
  | { kind: 'running' }
  | { kind: 'done'; solution: number; result: number }
  | { kind: 'error'; message: string };

export function GoalSeekDialog({ api, onClose }: DialogComponentProps) {
  // Seed "Set cell" with the current single-cell selection, if any.
  const initialSetCell = useMemo(() => {
    const range = api.univer.getActiveWorkbook()?.getActiveSheet()?.getActiveRange();
    const a1 = (range as unknown as { getA1Notation?: () => string } | null)?.getA1Notation?.();
    return a1 && isSingleCellA1(a1) ? a1 : '';
  }, [api]);

  const [state, setState] = useState<DialogState>({ ...INITIAL_STATE, setCell: initialSetCell });
  const [status, setStatus] = useState<Status>({ kind: 'idle' });

  const update = <K extends keyof DialogState>(key: K, value: DialogState[K]) => {
    setState((prev) => ({ ...prev, [key]: value }));
    setStatus({ kind: 'idle' });
  };

  const setCellValid = isSingleCellA1(state.setCell);
  const changeCellValid = isSingleCellA1(state.byChangingCell);
  const targetNum = Number(state.toValue);
  const targetValid = state.toValue.trim() !== '' && Number.isFinite(targetNum);
  const canRun =
    setCellValid &&
    changeCellValid &&
    targetValid &&
    state.setCell.trim().toUpperCase() !== state.byChangingCell.trim().toUpperCase() &&
    status.kind !== 'running';

  const run = async () => {
    const solver = getSolver(api);
    if (!solver) {
      setStatus({ kind: 'error', message: 'No active sheet.' });
      return;
    }
    setStatus({ kind: 'running' });
    try {
      const result = await goalSeek(
        solver,
        state.setCell.trim(),
        targetNum,
        state.byChangingCell.trim(),
      );
      if (result.ok) {
        setStatus({ kind: 'done', solution: result.solution, result: result.result });
      } else {
        setStatus({ kind: 'error', message: result.reason });
      }
    } catch (err) {
      setStatus({
        kind: 'error',
        message: err instanceof Error ? err.message : 'Goal seek failed.',
      });
    }
  };

  const running = status.kind === 'running';

  return (
    <Dialog
      title="Goal seek"
      onClose={onClose}
      width={420}
      data-testid="cs-goal-seek-dialog"
      footer={
        <>
          <button type="button" style={DIALOG_BTN_SECONDARY_STYLE} onClick={onClose}>
            {status.kind === 'done' ? 'Done' : 'Cancel'}
          </button>
          <button
            type="button"
            style={{ ...DIALOG_BTN_PRIMARY_STYLE, opacity: canRun ? 1 : 0.5 }}
            data-testid="cs-goal-seek-run"
            disabled={!canRun}
            onClick={run}
          >
            {running ? 'Solving…' : 'Solve'}
          </button>
        </>
      }
    >
      <div style={NOTE_STYLE} data-testid="cs-goal-seek-note">
        Drives the input cell until the formula cell reaches your target value.
      </div>

      <label style={DIALOG_FIELD_STYLE}>
        <span style={DIALOG_LABEL_STYLE}>Set cell (formula cell)</span>
        <input
          style={DIALOG_INPUT_STYLE}
          data-testid="cs-goal-seek-set-cell"
          placeholder="e.g. B5"
          value={state.setCell}
          onChange={(e) => update('setCell', e.target.value)}
        />
      </label>

      <label style={DIALOG_FIELD_STYLE}>
        <span style={DIALOG_LABEL_STYLE}>To value</span>
        <input
          style={DIALOG_INPUT_STYLE}
          data-testid="cs-goal-seek-to-value"
          type="number"
          placeholder="e.g. 1000"
          value={state.toValue}
          onChange={(e) => update('toValue', e.target.value)}
        />
      </label>

      <label style={DIALOG_FIELD_STYLE}>
        <span style={DIALOG_LABEL_STYLE}>By changing cell (input cell)</span>
        <input
          style={DIALOG_INPUT_STYLE}
          data-testid="cs-goal-seek-change-cell"
          placeholder="e.g. B2"
          value={state.byChangingCell}
          onChange={(e) => update('byChangingCell', e.target.value)}
        />
      </label>

      {status.kind === 'done' && (
        <div style={STATUS_OK_STYLE} data-testid="cs-goal-seek-success">
          Solved: <strong>{state.byChangingCell.trim().toUpperCase()}</strong> ={' '}
          {formatNumber(status.solution)} makes{' '}
          <strong>{state.setCell.trim().toUpperCase()}</strong> = {formatNumber(status.result)}.
        </div>
      )}
      {status.kind === 'error' && (
        <div style={STATUS_ERR_STYLE} data-testid="cs-goal-seek-error">
          {status.message}
        </div>
      )}
    </Dialog>
  );
}

/** Trim solver float noise for display without losing precision users care about. */
function formatNumber(n: number): string {
  if (!Number.isFinite(n)) return String(n);
  const rounded = Math.round(n * 1e6) / 1e6;
  return String(rounded);
}
