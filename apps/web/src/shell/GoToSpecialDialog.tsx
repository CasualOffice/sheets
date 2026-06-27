import { useState } from 'react';
import type { FUniver } from '@univerjs/core/facade';
import { Dialog } from './Dialog';
import { computeGoToSpecial, type CellMatrix, type GoToCriterion } from './go-to-special';

/**
 * Go To Special (Excel's F5 → Special). Select every cell on the active sheet
 * matching a criterion — constants, formulas, blanks, the current region, or
 * the last cell of the used range. The selection maths lives in
 * `go-to-special.ts`; this dialog reads the active sheet's cell matrix, runs it,
 * and applies the resulting multi-range selection through the same
 * `set-selections` operation the grid uses.
 */

type Props = {
  api: FUniver;
  onClose: () => void;
};

const CHOICES: Array<{ id: GoToCriterion; label: string; hint: string }> = [
  { id: 'constants', label: 'Constants', hint: 'Cells with a typed value (not formulas)' },
  { id: 'formulas', label: 'Formulas', hint: 'Cells containing a formula' },
  { id: 'blanks', label: 'Blanks', hint: 'Empty cells within the used range' },
  { id: 'currentRegion', label: 'Current region', hint: 'The block around the active cell' },
  { id: 'lastCell', label: 'Last cell', hint: 'Bottom-right of the used range' },
];

const SET_SELECTIONS_OP_ID = 'sheet.operation.set-selections';

export function GoToSpecialDialog({ api, onClose }: Props) {
  const [criterion, setCriterion] = useState<GoToCriterion>('constants');
  const [notice, setNotice] = useState<string | null>(null);

  const apply = async () => {
    setNotice(null);
    const wb = api.getActiveWorkbook();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws = wb?.getActiveSheet() as any;
    if (!wb || !ws) return;

    const subUnitId: string = ws.getSheetId();
    const unitId: string = wb.getId();
    const activeRange = ws.getActiveRange?.();
    const active = {
      row: activeRange?.getRow?.() ?? 0,
      column: activeRange?.getColumn?.() ?? 0,
    };

    // The active sheet's cell matrix from the snapshot (v + f per cell).
    const snapshot = wb.save();
    const cellData = (snapshot?.sheets?.[subUnitId]?.cellData ?? {}) as CellMatrix;

    const { ranges } = computeGoToSpecial(cellData, active, criterion);
    if (ranges.length === 0) {
      setNotice('No cells were found.');
      return;
    }

    const params = {
      unitId,
      subUnitId,
      selections: ranges.map((r, i) => ({
        range: {
          startRow: r.startRow,
          startColumn: r.startColumn,
          endRow: r.endRow,
          endColumn: r.endColumn,
          rangeType: 0, // RANGE_TYPE.NORMAL
        },
        // The first matched range carries the active cell (primary).
        primary:
          i === 0
            ? {
                startRow: r.startRow,
                startColumn: r.startColumn,
                endRow: r.startRow,
                endColumn: r.startColumn,
                actualRow: r.startRow,
                actualColumn: r.startColumn,
                isMerged: false,
                isMergedMainCell: false,
              }
            : null,
        style: null,
      })),
    };

    try {
      await api.executeCommand(SET_SELECTIONS_OP_ID, params);
      onClose();
    } catch (err) {
      console.warn('[go-to-special] failed to apply selection', err);
      setNotice('Could not apply the selection — see console.');
    }
  };

  return (
    <Dialog
      title="Go To Special"
      onClose={onClose}
      data-testid="go-to-special-dialog"
      footer={
        <>
          <button
            type="button"
            className="btn-secondary"
            data-testid="go-to-special-cancel"
            onClick={onClose}
          >
            Cancel
          </button>
          <button
            type="button"
            className="btn-primary"
            data-testid="go-to-special-ok"
            onClick={() => void apply()}
          >
            OK
          </button>
        </>
      }
    >
      <div role="radiogroup" aria-label="Select" style={{ display: 'grid', gap: 4 }}>
        {CHOICES.map((c) => (
          <label
            key={c.id}
            style={{
              display: 'flex',
              alignItems: 'baseline',
              gap: 8,
              padding: '6px 8px',
              borderRadius: 6,
              cursor: 'pointer',
              background: criterion === c.id ? 'var(--cs-chrome-active, #e6f3f7)' : 'transparent',
            }}
          >
            <input
              type="radio"
              name="go-to-special"
              value={c.id}
              checked={criterion === c.id}
              data-testid={`go-to-special-${c.id}`}
              onChange={() => {
                setCriterion(c.id);
                setNotice(null);
              }}
            />
            <span style={{ display: 'grid', gap: 1 }}>
              <span style={{ fontSize: 13 }}>{c.label}</span>
              <span style={{ fontSize: 11, color: 'var(--cs-chrome-muted, #8a8886)' }}>
                {c.hint}
              </span>
            </span>
          </label>
        ))}
      </div>
      {notice && (
        <div
          data-testid="go-to-special-notice"
          style={{ marginTop: 10, fontSize: 12, color: 'var(--cs-chrome-muted, #8a8886)' }}
        >
          {notice}
        </div>
      )}
    </Dialog>
  );
}
