import { useState } from 'react';
import type { FUniver } from '@univerjs/core/facade';
import { Dialog } from './Dialog';
import { matchAdvancedFilter, type FilterValue } from './advanced-filter';

/**
 * Advanced Filter (Excel's Data → Advanced), "copy to another location" mode.
 * Reads the list range + a criteria range (header of field names, rows ORed,
 * columns ANDed) and writes the header + matching rows to a destination. The
 * matching grammar lives in `advanced-filter.ts`; this dialog just resolves the
 * three ranges and writes the result with one `setValues`.
 */

type Props = {
  api: FUniver;
  onClose: () => void;
};

export function AdvancedFilterDialog({ api, onClose }: Props) {
  const wb = api.getActiveWorkbook();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const ws = wb?.getActiveSheet() as any;
  const activeA1 = ws?.getActiveRange?.()?.getA1Notation?.() ?? '';

  const [listRef, setListRef] = useState<string>(activeA1);
  const [critRef, setCritRef] = useState<string>('');
  const [destRef, setDestRef] = useState<string>('');
  const [result, setResult] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);

  const getRange = (a1: string) => {
    try {
      const r = ws?.getRange?.(a1.trim());
      return r && r.getWidth?.() > 0 ? r : null;
    } catch {
      return null;
    }
  };

  const apply = () => {
    setError(null);
    setResult(null);
    if (!ws) return;
    const list = getRange(listRef);
    const crit = getRange(critRef);
    if (!list) return setError('Enter a valid list range (e.g. A1:C20).');
    if (!crit) return setError('Enter a valid criteria range (e.g. E1:F2).');
    const dest = getRange(destRef);
    if (!dest) return setError('Enter a destination cell to copy to (e.g. H1).');

    const listVals = list.getValues() as FilterValue[][];
    const critVals = crit.getValues() as FilterValue[][];
    if (listVals.length < 2) return setError('The list range needs a header row and data.');

    const listHeader = listVals[0];
    const listRows = listVals.slice(1);
    const critHeader = critVals[0] ?? [];
    const critRows = critVals.slice(1);

    const idx = matchAdvancedFilter({ listHeader, listRows, critHeader, critRows });
    const output: FilterValue[][] = [listHeader, ...idx.map((i) => listRows[i])];

    const destRow = dest.getRow();
    const destCol = dest.getColumn();
    const width = listHeader.length;
    const target = ws.getRange(destRow, destCol, output.length, width);
    target.setValues(output as never);
    setResult(`${idx.length} ${idx.length === 1 ? 'row' : 'rows'} copied.`);
  };

  const field = (
    label: string,
    value: string,
    set: (v: string) => void,
    testid: string,
    ph: string,
  ) => (
    <label style={{ display: 'grid', gap: 4, fontSize: 13 }}>
      {label}
      <input
        type="text"
        data-testid={testid}
        value={value}
        spellCheck={false}
        placeholder={ph}
        onChange={(e) => set(e.target.value.toUpperCase())}
      />
    </label>
  );

  return (
    <Dialog
      title="Advanced Filter"
      onClose={onClose}
      data-testid="advanced-filter-dialog"
      footer={
        <>
          <button
            type="button"
            className="btn-secondary"
            data-testid="advanced-filter-close"
            onClick={onClose}
          >
            {result ? 'Close' : 'Cancel'}
          </button>
          <button
            type="button"
            className="btn-primary"
            data-testid="advanced-filter-ok"
            onClick={apply}
          >
            OK
          </button>
        </>
      }
    >
      <div style={{ display: 'grid', gap: 10 }}>
        <div style={{ fontSize: 12, color: 'var(--cs-chrome-muted, #8a8886)' }}>
          Copies rows matching the criteria range to another location. The criteria range has a
          header of field names; rows are ORed, columns ANDed.
        </div>
        {field('List range', listRef, setListRef, 'advanced-filter-list', 'A1:C20')}
        {field('Criteria range', critRef, setCritRef, 'advanced-filter-criteria', 'E1:F2')}
        {field('Copy to', destRef, setDestRef, 'advanced-filter-dest', 'H1')}
        {error && (
          <div data-testid="advanced-filter-error" style={{ fontSize: 12, color: '#b00020' }}>
            {error}
          </div>
        )}
        {result && (
          <div data-testid="advanced-filter-result" style={{ fontSize: 13 }}>
            {result}
          </div>
        )}
      </div>
    </Dialog>
  );
}
