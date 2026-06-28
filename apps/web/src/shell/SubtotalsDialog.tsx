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

import { useMemo, useState } from 'react';
import type { FUniver } from '@univerjs/core/facade';
import { Dialog } from './Dialog';
import { computeSubtotals, type SubtotalRow } from './subtotals';

/**
 * Subtotals (Excel's Data → Subtotal). Inserts a SUBTOTAL row at each change in
 * the grouping column plus a grand total. The layout + formula maths lives in
 * `subtotals.ts`; this dialog gathers the options, makes room with one
 * `insertRows`, and writes the new block back with one `setValues`. The list is
 * assumed to have a header row and to be sorted by the grouping column (as Excel
 * expects) — grouping is by consecutive runs.
 */

type Props = {
  api: FUniver;
  onClose: () => void;
};

const FUNCTIONS: Array<{ code: number; label: string }> = [
  { code: 9, label: 'Sum' },
  { code: 3, label: 'Count' },
  { code: 1, label: 'Average' },
  { code: 4, label: 'Max' },
  { code: 5, label: 'Min' },
  { code: 6, label: 'Product' },
  { code: 2, label: 'Count Numbers' },
];

function colToLetters(col: number): string {
  let s = '';
  let n = col;
  do {
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return s;
}

const headerText = (v: unknown): string =>
  v == null
    ? ''
    : typeof v === 'object' && 'v' in (v as Record<string, unknown>)
      ? String((v as { v: unknown }).v ?? '')
      : String(v);

export function SubtotalsDialog({ api, onClose }: Props) {
  const wb = api.getActiveWorkbook();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const ws = wb?.getActiveSheet() as any;
  const range = ws?.getActiveRange?.();

  const width: number = range?.getWidth?.() ?? 0;
  const height: number = range?.getHeight?.() ?? 0;
  const startRow: number = range?.getRow?.() ?? 0;
  const startCol: number = range?.getColumn?.() ?? 0;
  const usable = !!range && width >= 1 && height >= 2;

  const headers: string[] = useMemo(() => {
    if (!range) return [];
    const first = (range.getValues?.()?.[0] ?? []) as unknown[];
    return Array.from({ length: width }, (_, i) => {
      const h = headerText(first[i]).trim();
      return h !== '' ? h : `Column ${colToLetters(startCol + i)}`;
    });
  }, [range, width, startCol]);

  const [groupCol, setGroupCol] = useState(0);
  const [fnCode, setFnCode] = useState(9);
  const [subtotalCols, setSubtotalCols] = useState<boolean[]>(() =>
    Array.from({ length: width }, (_, i) => i === width - 1),
  );
  const [notice, setNotice] = useState<string | null>(null);

  const apply = () => {
    if (!usable || !range) {
      setNotice('Select a list with a header row and at least one data row.');
      return;
    }
    const cols = subtotalCols.map((on, i) => (on ? i : -1)).filter((i) => i >= 0);
    if (cols.length === 0) {
      setNotice('Choose at least one column to subtotal.');
      return;
    }
    const rows = range.getCellDatas() as SubtotalRow[];
    const out = computeSubtotals({
      rows,
      groupCol,
      functionCode: fnCode,
      subtotalCols: cols,
      startRow,
      startCol,
    });
    if (out.insertedRows > 0) {
      // Make room below the list, then paint the interleaved block in one write.
      ws.insertRows(startRow + height, out.insertedRows);
    }
    const target = ws.getRange(startRow, startCol, out.rows.length, width);
    target.setValues(out.rows as never);
    onClose();
  };

  return (
    <Dialog
      title="Subtotal"
      onClose={onClose}
      data-testid="subtotals-dialog"
      footer={
        <>
          <button
            type="button"
            className="btn-secondary"
            data-testid="subtotals-cancel"
            onClick={onClose}
          >
            Cancel
          </button>
          <button
            type="button"
            className="btn-primary"
            data-testid="subtotals-ok"
            onClick={apply}
            disabled={!usable}
          >
            OK
          </button>
        </>
      }
    >
      {!usable ? (
        <div data-testid="subtotals-notice" style={{ fontSize: 13 }}>
          Select a list with a header row and at least one data row.
        </div>
      ) : (
        <div style={{ display: 'grid', gap: 10 }}>
          <label style={{ display: 'grid', gap: 4, fontSize: 13 }}>
            At each change in:
            <select
              data-testid="subtotals-group"
              value={groupCol}
              onChange={(e) => setGroupCol(Number(e.target.value))}
            >
              {headers.map((h, i) => (
                <option key={i} value={i}>
                  {h}
                </option>
              ))}
            </select>
          </label>
          <label style={{ display: 'grid', gap: 4, fontSize: 13 }}>
            Use function:
            <select
              data-testid="subtotals-function"
              value={fnCode}
              onChange={(e) => setFnCode(Number(e.target.value))}
            >
              {FUNCTIONS.map((f) => (
                <option key={f.code} value={f.code}>
                  {f.label}
                </option>
              ))}
            </select>
          </label>
          <div style={{ fontSize: 13 }}>
            <div style={{ marginBottom: 4 }}>Add subtotal to:</div>
            <div
              style={{
                maxHeight: 180,
                overflowY: 'auto',
                border: '1px solid var(--cs-chrome-border, #e6e9ee)',
                borderRadius: 6,
                padding: 4,
              }}
            >
              {headers.map((h, i) => (
                <label
                  key={i}
                  style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '3px 6px' }}
                >
                  <input
                    type="checkbox"
                    data-testid={`subtotals-col-${i}`}
                    checked={subtotalCols[i] ?? false}
                    onChange={(e) =>
                      setSubtotalCols((cur) => {
                        const next = [...cur];
                        next[i] = e.target.checked;
                        return next;
                      })
                    }
                  />
                  {h}
                </label>
              ))}
            </div>
          </div>
          {notice && (
            <div
              data-testid="subtotals-warn"
              style={{ fontSize: 12, color: 'var(--cs-chrome-muted, #8a8886)' }}
            >
              {notice}
            </div>
          )}
        </div>
      )}
    </Dialog>
  );
}
