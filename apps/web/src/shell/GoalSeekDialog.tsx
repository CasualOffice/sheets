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
import { runGoalSeek, type GoalSeekResult } from './goal-seek';

/**
 * Excel-style Goal Seek dialog: drive one cell to a target value by
 * varying another. Three fields:
 *
 *   - Set cell:  the goal cell (A1 notation)
 *   - To value:  the target numeric value
 *   - By changing: the input cell to vary
 *
 * Convergence + result reporting is handled in `goal-seek.ts`. The
 * dialog stays modal until the user clicks OK or Cancel; we surface
 * the final value + iteration count so the user knows whether the
 * solver succeeded.
 */

type Props = {
  api: FUniver;
  onClose: () => void;
};

export function GoalSeekDialog({ api, onClose }: Props) {
  const initialActive = readActiveCellA1(api);
  const [goalRef, setGoalRef] = useState(initialActive ?? '');
  const [target, setTarget] = useState('0');
  const [inputRef, setInputRef] = useState('');
  const [running, setRunning] = useState(false);
  const [result, setResult] = useState<GoalSeekResult | null>(null);
  const [error, setError] = useState<string | null>(null);

  const run = async () => {
    setError(null);
    const goal = parseA1(goalRef);
    const input = parseA1(inputRef);
    const targetNum = Number(target);
    if (!goal) {
      setError('Set cell must be a cell reference like A1.');
      return;
    }
    if (!input) {
      setError('By changing cell must be a cell reference like B5.');
      return;
    }
    if (!Number.isFinite(targetNum)) {
      setError('To value must be a number.');
      return;
    }
    setRunning(true);
    try {
      const out = await runGoalSeek(api, {
        goalCell: goal,
        targetValue: targetNum,
        inputCell: input,
      });
      setResult(out);
    } catch (err) {
      console.warn('[goal-seek] runtime error', err);
      setError('Goal Seek hit a runtime error — see console.');
    } finally {
      setRunning(false);
    }
  };

  return (
    <Dialog
      title="Goal Seek"
      onClose={onClose}
      data-testid="goal-seek-dialog"
      footer={
        <>
          <button
            type="button"
            className="btn-secondary"
            data-testid="goal-seek-cancel"
            onClick={onClose}
          >
            Close
          </button>
          <button
            type="button"
            className="btn-primary"
            data-testid="goal-seek-run"
            onClick={() => void run()}
            disabled={running}
          >
            {running ? 'Solving…' : 'OK'}
          </button>
        </>
      }
    >
      <div className="goal-seek">
        <label className="goal-seek__field">
          <span>Set cell</span>
          <input
            type="text"
            data-testid="goal-seek-goal-input"
            value={goalRef}
            onChange={(e) => setGoalRef(e.target.value.toUpperCase())}
            placeholder="A1"
            spellCheck={false}
          />
        </label>
        <label className="goal-seek__field">
          <span>To value</span>
          <input
            type="text"
            inputMode="decimal"
            data-testid="goal-seek-target-input"
            value={target}
            onChange={(e) => setTarget(e.target.value)}
            placeholder="100"
          />
        </label>
        <label className="goal-seek__field">
          <span>By changing cell</span>
          <input
            type="text"
            data-testid="goal-seek-input-input"
            value={inputRef}
            onChange={(e) => setInputRef(e.target.value.toUpperCase())}
            placeholder="B5"
            spellCheck={false}
          />
        </label>
        {error && (
          <div className="goal-seek__error" data-testid="goal-seek-error">
            {error}
          </div>
        )}
        {result && (
          <div
            className={`goal-seek__result goal-seek__result--${result.ok ? 'ok' : 'fail'}`}
            data-testid="goal-seek-result"
          >
            {result.ok ? (
              <>
                Converged in {result.iterations} {result.iterations === 1 ? 'step' : 'steps'}.
                Set <code>{inputRef}</code> to <strong>{formatFinal(result.finalValue)}</strong>.
              </>
            ) : (
              <>
                {result.reason} Best candidate after {result.iterations} steps:{' '}
                <strong>{result.finalValue != null ? formatFinal(result.finalValue) : '—'}</strong>.
              </>
            )}
          </div>
        )}
      </div>
    </Dialog>
  );
}

function formatFinal(n: number): string {
  // Trim trailing zeros while keeping precision for users running the
  // solver with small tolerances. 6 sig figs is enough for almost
  // every spreadsheet use case.
  return Number(n.toPrecision(6)).toString();
}

function parseA1(s: string): { row: number; col: number } | null {
  const m = /^\$?([A-Z]+)\$?(\d+)$/.exec(s.trim());
  if (!m) return null;
  const letters = m[1];
  let col = 0;
  for (let i = 0; i < letters.length; i += 1) {
    col = col * 26 + (letters.charCodeAt(i) - 64);
  }
  return { col: col - 1, row: parseInt(m[2], 10) - 1 };
}

function readActiveCellA1(api: FUniver): string | null {
  const wb = api.getActiveWorkbook();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const ws = wb?.getActiveSheet() as any;
  const range = ws?.getActiveRange?.();
  if (!range) return null;
  const col = range.getColumn();
  const row = range.getRow();
  let n = col;
  let letters = '';
  while (n >= 0) {
    letters = String.fromCharCode(65 + (n % 26)) + letters;
    n = Math.floor(n / 26) - 1;
  }
  return `${letters}${row + 1}`;
}
