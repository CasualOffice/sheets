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
import { Dialog } from '../shell/Dialog';
import type { SparklineType } from './types';

/**
 * Excel-style Insert Sparkline dialog. Three inputs:
 *
 *   - Data range (A1:F1 etc.) — values to plot.
 *   - Location (single cell) — where the mini-chart renders.
 *   - Type — line / column / win-loss.
 *
 * v1 keeps the source + anchor on the same sheet; cross-sheet
 * sparklines are a follow-up gated on demand.
 */

type Props = {
  api: FUniver;
  defaultSourceA1?: string;
  defaultAnchorA1?: string;
  onCancel: () => void;
  onConfirm: (params: {
    type: SparklineType;
    source: { startRow: number; endRow: number; startColumn: number; endColumn: number };
    anchor: { row: number; col: number };
  }) => void;
};

export function InsertSparklineDialog({
  api,
  defaultSourceA1,
  defaultAnchorA1,
  onCancel,
  onConfirm,
}: Props) {
  const [source, setSource] = useState(defaultSourceA1 ?? '');
  const [anchor, setAnchor] = useState(defaultAnchorA1 ?? '');
  const [type, setType] = useState<SparklineType>('line');
  const [error, setError] = useState<string | null>(null);

  const submit = () => {
    void api; // kept for future cross-sheet validation
    const src = parseRange(source);
    const anc = parseSingleCell(anchor);
    if (!src) {
      setError('Data range must be a range like A1:F1.');
      return;
    }
    if (!anc) {
      setError('Location must be a single cell like G1.');
      return;
    }
    onConfirm({ type, source: src, anchor: anc });
  };

  return (
    <Dialog
      title="Insert Sparkline"
      onClose={onCancel}
      data-testid="insert-sparkline-dialog"
      footer={
        <>
          <button
            type="button"
            className="btn-secondary"
            data-testid="insert-sparkline-cancel"
            onClick={onCancel}
          >
            Cancel
          </button>
          <button
            type="button"
            className="btn-primary"
            data-testid="insert-sparkline-ok"
            onClick={submit}
          >
            OK
          </button>
        </>
      }
    >
      <div className="insert-sparkline">
        <label className="insert-sparkline__field">
          <span>Data range</span>
          <input
            type="text"
            data-testid="insert-sparkline-source"
            value={source}
            placeholder="A1:F1"
            onChange={(e) => setSource(e.target.value.toUpperCase())}
            spellCheck={false}
          />
        </label>
        <label className="insert-sparkline__field">
          <span>Location</span>
          <input
            type="text"
            data-testid="insert-sparkline-anchor"
            value={anchor}
            placeholder="G1"
            onChange={(e) => setAnchor(e.target.value.toUpperCase())}
            spellCheck={false}
          />
        </label>
        <fieldset className="insert-sparkline__type" role="radiogroup" aria-label="Sparkline type">
          <legend>Type</legend>
          {(['line', 'column', 'win-loss'] as SparklineType[]).map((t) => (
            <label
              key={t}
              className={`insert-sparkline__type-opt${type === t ? ' insert-sparkline__type-opt--active' : ''}`}
              data-testid={`insert-sparkline-type-${t}`}
            >
              <input
                type="radio"
                name="spark-type"
                value={t}
                checked={type === t}
                onChange={() => setType(t)}
              />
              <span>{t === 'win-loss' ? 'Win / Loss' : t.charAt(0).toUpperCase() + t.slice(1)}</span>
            </label>
          ))}
        </fieldset>
        {error && (
          <div className="insert-sparkline__error" data-testid="insert-sparkline-error">
            {error}
          </div>
        )}
      </div>
    </Dialog>
  );
}

function parseRange(
  s: string,
): { startRow: number; endRow: number; startColumn: number; endColumn: number } | null {
  const trimmed = s.trim().toUpperCase();
  const m = /^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/.exec(trimmed);
  if (!m) return null;
  const c1 = colLettersToIndex(m[1]);
  const r1 = parseInt(m[2], 10) - 1;
  const c2 = m[3] ? colLettersToIndex(m[3]) : c1;
  const r2 = m[4] ? parseInt(m[4], 10) - 1 : r1;
  return {
    startRow: Math.min(r1, r2),
    endRow: Math.max(r1, r2),
    startColumn: Math.min(c1, c2),
    endColumn: Math.max(c1, c2),
  };
}

function parseSingleCell(s: string): { row: number; col: number } | null {
  const trimmed = s.trim().toUpperCase();
  const m = /^\$?([A-Z]+)\$?(\d+)$/.exec(trimmed);
  if (!m) return null;
  return { col: colLettersToIndex(m[1]), row: parseInt(m[2], 10) - 1 };
}

function colLettersToIndex(letters: string): number {
  let n = 0;
  for (let i = 0; i < letters.length; i += 1) {
    n = n * 26 + (letters.charCodeAt(i) - 64);
  }
  return n - 1;
}
