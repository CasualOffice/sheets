import { useMemo, useState } from 'react';
import type { FUniver } from '@univerjs/core/facade';
import { Dialog } from './Dialog';
import { dedupeRows, type CellRow } from './dedupe';

/**
 * Remove Duplicates (Excel's Data → Remove Duplicates). Drops rows that
 * duplicate an earlier row across the chosen columns, keeping the first
 * occurrence, then compacts the survivors to the top of the selection
 * (preserving value + formula + style) and clears the freed rows. The dedupe
 * maths lives in `dedupe.ts`; this dialog reads the selection's cell data, runs
 * it, writes back with one `setValues`, and reports the counts like Excel.
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

export function RemoveDuplicatesDialog({ api, onClose }: Props) {
  const wb = api.getActiveWorkbook();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const ws = wb?.getActiveSheet() as any;
  const range = ws?.getActiveRange?.();

  const width: number = range?.getWidth?.() ?? 0;
  const startCol: number = range?.getColumn?.() ?? 0;
  // First row's display values, for header-mode column labels.
  const firstRow: unknown[] = useMemo(() => (range ? (range.getValues()?.[0] ?? []) : []), [range]);

  const [hasHeaders, setHasHeaders] = useState(true);
  const [selected, setSelected] = useState<boolean[]>(() => new Array(width).fill(true));
  const [result, setResult] = useState<{ removed: number; remaining: number } | null>(null);

  const columnLabel = (i: number): string => {
    const letter = colToLetters(startCol + i);
    if (hasHeaders) {
      const h = stringifyHeader(firstRow[i]).trim();
      return h !== '' ? h : `Column ${letter}`;
    }
    return `Column ${letter}`;
  };

  const apply = () => {
    if (!range) return;
    const compareColumns = selected.map((sel, i) => (sel ? i : -1)).filter((i) => i >= 0);
    const rows = range.getCellDatas() as CellRow[];
    const out = dedupeRows({ rows, hasHeaders, compareColumns });
    // Single atomic, undoable write — preserves v/f/s on the kept rows.
    range.setValues(out.rows as never);
    setResult({ removed: out.removed, remaining: out.remaining });
  };

  const allChecked = selected.every(Boolean);
  const toggleAll = () => setSelected(new Array(width).fill(!allChecked));

  return (
    <Dialog
      title="Remove Duplicates"
      onClose={onClose}
      data-testid="remove-duplicates-dialog"
      footer={
        result ? (
          <button
            type="button"
            className="btn-primary"
            data-testid="remove-duplicates-close"
            onClick={onClose}
          >
            Close
          </button>
        ) : (
          <>
            <button
              type="button"
              className="btn-secondary"
              data-testid="remove-duplicates-cancel"
              onClick={onClose}
            >
              Cancel
            </button>
            <button
              type="button"
              className="btn-primary"
              data-testid="remove-duplicates-ok"
              onClick={apply}
              disabled={width === 0}
            >
              OK
            </button>
          </>
        )
      }
    >
      {result ? (
        <div data-testid="remove-duplicates-result" style={{ fontSize: 13, lineHeight: 1.5 }}>
          {result.removed === 0 ? (
            <>No duplicate values found.</>
          ) : (
            <>
              <strong>{result.removed}</strong> duplicate{' '}
              {result.removed === 1 ? 'value' : 'values'} found and removed;{' '}
              <strong>{result.remaining}</strong> unique{' '}
              {result.remaining === 1 ? 'value remains' : 'values remain'}.
            </>
          )}
        </div>
      ) : (
        <>
          <label
            style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: 13, marginBottom: 8 }}
          >
            <input
              type="checkbox"
              data-testid="remove-duplicates-headers"
              checked={hasHeaders}
              onChange={(e) => setHasHeaders(e.target.checked)}
            />
            My data has headers
          </label>
          <div style={{ fontSize: 12, color: 'var(--cs-chrome-muted, #8a8886)', marginBottom: 4 }}>
            Columns to compare
          </div>
          <div
            style={{
              maxHeight: 220,
              overflowY: 'auto',
              border: '1px solid var(--cs-chrome-border, #e6e9ee)',
              borderRadius: 6,
              padding: 4,
            }}
          >
            <label
              style={{
                display: 'flex',
                alignItems: 'center',
                gap: 8,
                padding: '4px 6px',
                fontSize: 13,
                fontWeight: 600,
              }}
            >
              <input
                type="checkbox"
                data-testid="remove-duplicates-all"
                checked={allChecked}
                onChange={toggleAll}
              />
              Select all
            </label>
            {Array.from({ length: width }, (_, i) => (
              <label
                key={i}
                style={{
                  display: 'flex',
                  alignItems: 'center',
                  gap: 8,
                  padding: '4px 6px',
                  fontSize: 13,
                }}
              >
                <input
                  type="checkbox"
                  data-testid={`remove-duplicates-col-${i}`}
                  checked={selected[i] ?? false}
                  onChange={(e) =>
                    setSelected((cur) => {
                      const next = [...cur];
                      next[i] = e.target.checked;
                      return next;
                    })
                  }
                />
                {columnLabel(i)}
              </label>
            ))}
          </div>
        </>
      )}
    </Dialog>
  );
}
