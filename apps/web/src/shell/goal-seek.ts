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

import type { FUniver } from '@univerjs/core/facade';
import { ICommandService, type ICommandInfo } from '@univerjs/core';

/**
 * Excel's Goal Seek — iteratively solve `f(x) = target` by varying a
 * single input cell. The "function" is whatever the workbook computes
 * for the goal cell when the input cell holds a candidate value, so
 * the solver works for any formula chain Univer can evaluate.
 *
 * Algorithm: combined bisection + secant. Bisection always converges
 * if the goal is bracketed; secant accelerates once we're close.
 * Convergence target: |f(x) - goal| ≤ TOL, or iteration cap reached.
 *
 * Side effect: the input cell is updated to the converged value (or
 * left at the best-found candidate if convergence failed). The
 * mutation goes through `set-range-values` like a normal user edit so
 * co-edit propagation, undo, and xlsx round-trip all work.
 */

const MAX_ITER = 100;
const TOL = 1e-6;
/** Multiplier when searching for an initial bracket. We start from a
 *  candidate value and expand outwards by ±STEP × scale until the
 *  goal function changes sign on at least one side. */
const STEP = 1;
const STEP_GROWTH = 2;
const MAX_BRACKET_TRIES = 20;
/** If the goal cell's value doesn't change after this many probes,
 *  the input cell is not in the formula's dependency tree — bail
 *  with a clear error instead of expanding to ±2^MAX_BRACKET_TRIES. */
const CONSTANT_BAILOUT = 3;

export type GoalSeekResult =
  | { ok: true; iterations: number; finalValue: number }
  | { ok: false; reason: string; iterations: number; finalValue: number | null };

export type GoalSeekParams = {
  goalCell: { row: number; col: number }; // The cell whose value we're driving
  targetValue: number;
  inputCell: { row: number; col: number }; // The cell we're varying
};

export async function runGoalSeek(
  api: FUniver,
  params: GoalSeekParams,
): Promise<GoalSeekResult> {
  const wb = api.getActiveWorkbook();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const sheet = wb?.getActiveSheet() as any;
  if (!wb || !sheet) {
    return { ok: false, reason: 'No active sheet', iterations: 0, finalValue: null };
  }
  const unitId: string = wb.getId();
  const subUnitId: string = sheet.getSheetId();

  // Formula evaluation runs in a Web Worker (see plugins.ts —
  // `notExecuteFormula: true` on the main-thread sheets plugin). That
  // means writing the input cell is synchronous, but the dependent
  // formula's recompute lands one round-trip later as a
  // `formula.mutation.set-formula-calculation-result` mutation. Without
  // waiting for it we'd read stale goal-cell values and the solver
  // never converges.
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const injector = (api as any)._injector as
    | { get: (token: unknown) => unknown }
    | undefined;
  if (!injector) {
    return { ok: false, reason: 'No injector', iterations: 0, finalValue: null };
  }
  const cmdSvc = injector.get(ICommandService) as {
    onMutationExecutedForCollab: (
      l: (info: ICommandInfo) => void,
    ) => { dispose: () => void };
  };

  const waitForRecalc = (): Promise<void> =>
    new Promise<void>((resolve) => {
      // Tight cap so probes that don't trigger a recompute (e.g. the
      // goal cell is a literal, not a formula) don't stall the solver.
      // 80 ms is comfortably more than the worker round-trip on a
      // typical workbook.
      const timer = setTimeout(() => {
        sub.dispose();
        resolve();
      }, 80);
      const sub = cmdSvc.onMutationExecutedForCollab((info) => {
        if (info.id === 'formula.mutation.set-formula-calculation-result') {
          clearTimeout(timer);
          sub.dispose();
          resolve();
        }
      });
    });

  const readGoal = (): number => {
    const data = sheet.getRange(params.goalCell.row, params.goalCell.col).getCellData();
    const raw = data?.v;
    const n = typeof raw === 'number' ? raw : Number(raw);
    return Number.isFinite(n) ? n : NaN;
  };

  /** Write a candidate into the input cell, wait for the worker to
   *  recompute dependents, then read back the goal cell. We race two
   *  signals: the `formula.mutation.set-formula-calculation-result`
   *  hook (preferred — instant) and a short value-change poll (the
   *  fallback when the goal cell is a literal, not a formula).        */
  const probe = async (candidate: number): Promise<number> => {
    const prev = readGoal();
    const wait = waitForRecalc();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    await (api as any).executeCommand('sheet.command.set-range-values', {
      unitId,
      subUnitId,
      value: { [params.inputCell.row]: { [params.inputCell.col]: { v: candidate } } },
    });
    await wait;
    // Poll briefly for the value-change as a backstop — the
    // mutation-based wait above sometimes resolves a tick before the
    // cell store has applied the new value.
    let curr = readGoal();
    let polls = 0;
    while ((curr === prev || Number.isNaN(curr)) && polls < 8) {
      await new Promise((r) => setTimeout(r, 15));
      curr = readGoal();
      polls += 1;
    }
    return curr;
  };

  // Read the starting candidate from the input cell (or default to 0
  // if blank). Excel uses the cell's current value as the seed.
  const startData = sheet.getRange(params.inputCell.row, params.inputCell.col).getCellData();
  const startRaw = startData?.v;
  const startN = typeof startRaw === 'number' ? startRaw : Number(startRaw);
  let x0 = Number.isFinite(startN) ? startN : 0;

  // Evaluate at the seed.
  let f0 = (await probe(x0)) - params.targetValue;
  if (Math.abs(f0) <= TOL) {
    return { ok: true, iterations: 0, finalValue: x0 };
  }

  // Find a bracket by expanding outwards. We search both directions
  // from the seed, doubling each step, until we get a sign change.
  let x1 = x0 + STEP;
  let f1 = (await probe(x1)) - params.targetValue;
  let bracketTry = 0;
  let unchangedRuns = 0;
  // Early-exit "constant function" detection: if the goal value
  // returns identical for the seed and the first few candidates, the
  // input clearly doesn't drive the goal and the bisection would just
  // burn iterations.
  if (f0 === f1) unchangedRuns += 1;
  while (Math.sign(f0) === Math.sign(f1) && bracketTry < MAX_BRACKET_TRIES) {
    bracketTry += 1;
    // Reflect the search direction every other step so we sweep both
    // sides of the seed.
    const step = STEP * Math.pow(STEP_GROWTH, bracketTry) * (bracketTry % 2 === 0 ? 1 : -1);
    x1 = x0 + step;
    const newF = (await probe(x1)) - params.targetValue;
    if (newF === f1) {
      unchangedRuns += 1;
    } else {
      unchangedRuns = 0;
    }
    f1 = newF;
    if (Math.abs(f1) <= TOL) {
      return { ok: true, iterations: bracketTry, finalValue: x1 };
    }
    if (unchangedRuns >= CONSTANT_BAILOUT) {
      return {
        ok: false,
        reason: 'Changing the input cell does not change the goal cell.',
        iterations: bracketTry,
        finalValue: x1,
      };
    }
  }

  if (Math.sign(f0) === Math.sign(f1)) {
    // No bracket found. Leave the input cell at the best candidate
    // we've seen (closer of x0 / x1 in |f|).
    const bestX = Math.abs(f0) < Math.abs(f1) ? x0 : x1;
    await probe(bestX);
    return {
      ok: false,
      reason: 'Could not find a bracket — function may not cross the target.',
      iterations: bracketTry,
      finalValue: bestX,
    };
  }

  // Bisection with secant acceleration. On each iteration we try the
  // secant step first; if it falls outside the bracket, fall back to
  // the midpoint.
  let iter = 0;
  let xMid = (x0 + x1) / 2;
  let fMid = f0;
  while (iter < MAX_ITER) {
    iter += 1;
    // Secant candidate.
    const denom = f1 - f0;
    const xSec = denom !== 0 ? x1 - f1 * ((x1 - x0) / denom) : (x0 + x1) / 2;
    xMid = xSec >= Math.min(x0, x1) && xSec <= Math.max(x0, x1) ? xSec : (x0 + x1) / 2;
    fMid = (await probe(xMid)) - params.targetValue;
    if (Math.abs(fMid) <= TOL) {
      return { ok: true, iterations: iter, finalValue: xMid };
    }
    // Narrow the bracket.
    if (Math.sign(fMid) === Math.sign(f0)) {
      x0 = xMid;
      f0 = fMid;
    } else {
      x1 = xMid;
      f1 = fMid;
    }
  }

  return {
    ok: false,
    reason: `Did not converge within ${MAX_ITER} iterations.`,
    iterations: iter,
    finalValue: xMid,
  };
}
