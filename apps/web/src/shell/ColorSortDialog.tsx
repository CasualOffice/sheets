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
import { colorMultiSort, distinctColors, type SortRow } from './color-sort';

/**
 * Sort by colour (Excel's Data → Sort → Sort On: Cell Color / Font Color).
 * The selection is sorted so rows whose key-column cell carries the chosen
 * background — or font — colour move to the top (or bottom); the rest keep
 * their order. Maths lives in `color-sort.ts`; this reads cell colours via
 * `getBackgrounds()` and font colours from `getCellStyles()`, the cells via
 * `getCellDatas()`, then writes back with one `setValues`.
 */

type Props = {
  api: FUniver;
  onClose: () => void;
};

function colToLetters(col: number): string {
  let s = '';
  let n = col;
  do {
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return s;
}

const stringifyHeader = (v: unknown): string => {
  if (v == null) return '';
  if (typeof v === 'object' && 'v' in (v as Record<string, unknown>)) {
    return String((v as { v: unknown }).v ?? '');
  }
  return String(v);
};

type SortOn = 'cell' | 'font';

// Univer reports unfilled cells as white and unset fonts as black/empty.
const NO_FILL = new Set(['#ffffff', '#fff', 'rgb(255,255,255)', '']);
const AUTO_FONT = new Set(['#000000', '#000', 'rgb(0,0,0)', '']);
const colorLabel = (c: string, on: SortOn): string => {
  if (on === 'cell') return NO_FILL.has(c.toLowerCase()) ? 'No fill' : c;
  return AUTO_FONT.has(c.toLowerCase()) ? 'Automatic' : c;
};

export function ColorSortDialog({ api, onClose }: Props) {
  const wb = api.getActiveWorkbook();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const ws = wb?.getActiveSheet() as any;
  const range = ws?.getActiveRange?.();

  const width: number = range?.getWidth?.() ?? 0;
  const startCol: number = range?.getColumn?.() ?? 0;
  const firstRow: unknown[] = useMemo(() => (range ? (range.getValues()?.[0] ?? []) : []), [range]);
  // Full colour matrices of the selection, read once on open. Backgrounds via
  // getBackgrounds(); font colours derived from the resolved cell styles.
  const backgrounds: string[][] = useMemo(
    () => (range?.getBackgrounds?.() as string[][] | undefined) ?? [],
    [range],
  );
  const fontColors: string[][] = useMemo(() => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const styles = (range?.getCellStyles?.() as any[][] | undefined) ?? [];
    return styles.map((row) => row.map((s) => s?.color?.rgb ?? ''));
  }, [range]);

  const [hasHeaders, setHasHeaders] = useState(true);
  const [keyColumn, setKeyColumn] = useState(0);
  const [sortOn, setSortOn] = useState<SortOn>('cell');
  // Colour priority, top-to-bottom (click order). Multiple colours → a
  // multi-level sort; one colour → that colour on top.
  const [order, setOrder] = useState<string[]>([]);
  const resetOrder = () => setOrder([]);
  const toggleColor = (c: string) =>
    setOrder((cur) => (cur.includes(c) ? cur.filter((x) => x !== c) : [...cur, c]));

  const columnLabel = (i: number): string => {
    const letter = colToLetters(startCol + i);
    if (hasHeaders) {
      const h = stringifyHeader(firstRow[i]).trim();
      return h !== '' ? h : `Column ${letter}`;
    }
    return `Column ${letter}`;
  };

  // The key column's colour per row (cell or font), then the distinct swatches.
  const colColors = useMemo(
    () => (sortOn === 'cell' ? backgrounds : fontColors).map((row) => row[keyColumn] ?? ''),
    [backgrounds, fontColors, sortOn, keyColumn],
  );
  const swatches = useMemo(() => distinctColors(colColors, hasHeaders), [colColors, hasHeaders]);

  const apply = () => {
    if (!range || order.length === 0) return;
    const rows = range.getCellDatas() as SortRow[];
    const out = colorMultiSort({ rows, colors: colColors, hasHeaders, order });
    range.setValues(out.rows as never);
    onClose();
  };

  return (
    <Dialog
      title="Sort by colour"
      onClose={onClose}
      data-testid="color-sort-dialog"
      footer={
        <>
          <button
            type="button"
            className="btn-secondary"
            data-testid="color-sort-cancel"
            onClick={onClose}
          >
            Cancel
          </button>
          <button
            type="button"
            className="btn-primary"
            data-testid="color-sort-ok"
            onClick={apply}
            disabled={width === 0 || order.length === 0}
          >
            Sort
          </button>
        </>
      }
    >
      <div
        style={{ display: 'flex', flexDirection: 'column', gap: 12, fontSize: 13, minWidth: 320 }}
      >
        <label style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
          <input
            type="checkbox"
            data-testid="color-sort-headers"
            checked={hasHeaders}
            onChange={(e) => {
              setHasHeaders(e.target.checked);
              resetOrder();
            }}
          />
          My data has headers
        </label>

        <label style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
          <span>Sort on</span>
          <select
            data-testid="color-sort-on"
            value={sortOn}
            onChange={(e) => {
              setSortOn(e.target.value as SortOn);
              resetOrder();
            }}
          >
            <option value="cell">Cell colour</option>
            <option value="font">Font colour</option>
          </select>
        </label>

        <label style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
          <span>Sort by column</span>
          <select
            data-testid="color-sort-column"
            value={keyColumn}
            onChange={(e) => {
              setKeyColumn(Number(e.target.value));
              resetOrder();
            }}
          >
            {Array.from({ length: width }, (_, i) => (
              <option key={i} value={i}>
                {columnLabel(i)}
              </option>
            ))}
          </select>
        </label>

        <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
          <span>Colour</span>
          <div
            data-testid="color-sort-swatches"
            style={{ display: 'flex', flexWrap: 'wrap', gap: 6 }}
          >
            {swatches.length === 0 ? (
              <span style={{ color: 'var(--color-text-secondary)' }}>No cells in selection</span>
            ) : (
              swatches.map((c, i) => {
                const sel = order.indexOf(c);
                const selected = sel >= 0;
                return (
                  <button
                    key={c || `c${i}`}
                    type="button"
                    data-testid={`color-sort-swatch-${i}`}
                    title={colorLabel(c, sortOn)}
                    aria-pressed={selected}
                    onClick={() => toggleColor(c)}
                    style={{
                      display: 'inline-flex',
                      alignItems: 'center',
                      gap: 6,
                      padding: '4px 8px',
                      borderRadius: 6,
                      border: selected
                        ? '2px solid var(--color-accent, #107c41)'
                        : '1px solid var(--color-divider)',
                      background: 'var(--color-surface)',
                      cursor: 'pointer',
                    }}
                  >
                    <span
                      aria-hidden
                      style={{
                        width: 16,
                        height: 16,
                        borderRadius: 3,
                        border: '1px solid var(--color-divider)',
                        background: c || 'transparent',
                      }}
                    />
                    <span>{colorLabel(c, sortOn)}</span>
                    {selected && (
                      <span
                        data-testid={`color-sort-rank-${i}`}
                        style={{
                          fontSize: 11,
                          fontWeight: 600,
                          color: 'var(--color-accent, #107c41)',
                        }}
                      >
                        {sel + 1}
                      </span>
                    )}
                  </button>
                );
              })
            )}
          </div>
          <p style={{ fontSize: 12, color: 'var(--color-text-secondary)', margin: 0 }}>
            Click colours in the order you want them — first goes on top. Unpicked colours stay
            below, in their current order.
          </p>
        </div>
      </div>
    </Dialog>
  );
}
