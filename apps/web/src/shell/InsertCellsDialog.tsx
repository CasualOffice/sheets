import { useState } from 'react';
import { Dialog } from './Dialog';

export type CellOpDirection =
  | 'shift-right'
  | 'shift-down'
  | 'entire-row'
  | 'entire-column';

type Props = {
  /** `'insert'` matches Excel's "Insert" dialog (Ctrl+Shift+= /
   *  Ctrl++); `'delete'` matches "Delete" (Ctrl+-). The four radio
   *  options + their labels are the same, only the verb differs. */
  mode: 'insert' | 'delete';
  onCancel: () => void;
  onConfirm: (dir: CellOpDirection) => void;
};

const OPTIONS: { id: CellOpDirection; insertLabel: string; deleteLabel: string }[] = [
  { id: 'shift-right',  insertLabel: 'Shift cells right',  deleteLabel: 'Shift cells left' },
  { id: 'shift-down',   insertLabel: 'Shift cells down',   deleteLabel: 'Shift cells up' },
  { id: 'entire-row',   insertLabel: 'Entire row',         deleteLabel: 'Entire row' },
  { id: 'entire-column',insertLabel: 'Entire column',      deleteLabel: 'Entire column' },
];

/**
 * Excel's Ctrl++ / Ctrl+- modal. Four radio options matching Excel's
 * "Insert" / "Delete" dialogs exactly — the shortcut goes straight to
 * this picker rather than just inserting/deleting a row blindly,
 * because the choice between row / column / shift-cells changes the
 * blast radius. Default is "Entire row" (Excel's default too).
 */
export function InsertCellsDialog({ mode, onCancel, onConfirm }: Props) {
  const [dir, setDir] = useState<CellOpDirection>('entire-row');
  const title = mode === 'insert' ? 'Insert' : 'Delete';
  return (
    <Dialog
      title={title}
      onClose={onCancel}
      data-testid={`${mode}-cells-dialog`}
      footer={
        <>
          <button
            type="button"
            className="btn-secondary"
            data-testid={`${mode}-cells-cancel`}
            onClick={onCancel}
          >
            Cancel
          </button>
          <button
            type="button"
            className="btn-primary"
            data-testid={`${mode}-cells-ok`}
            onClick={() => onConfirm(dir)}
          >
            OK
          </button>
        </>
      }
    >
      <div className="cells-op" role="radiogroup" aria-label={`${title} options`}>
        {OPTIONS.map((o) => {
          const label = mode === 'insert' ? o.insertLabel : o.deleteLabel;
          return (
            <label
              key={o.id}
              className={`cells-op__option${dir === o.id ? ' cells-op__option--active' : ''}`}
              data-testid={`${mode}-cells-${o.id}`}
            >
              <input
                type="radio"
                name="cells-op"
                value={o.id}
                checked={dir === o.id}
                onChange={() => setDir(o.id)}
              />
              <span>{label}</span>
            </label>
          );
        })}
      </div>
    </Dialog>
  );
}
