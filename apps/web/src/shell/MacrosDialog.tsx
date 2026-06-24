import { useState } from 'react';
import type { FUniver } from '@univerjs/core/facade';
import { Dialog } from './Dialog';
import {
  availableMacroLetters,
  deleteMacro,
  listMacros,
  runMacro,
  setMacroShortcut,
  type Macro,
} from '../sheets/macros';

/**
 * Manage Macros — Excel's Alt+F8 dialog: list saved macros, run or delete each.
 * Recording happens from the Macros menu; this is the management surface
 * (without it, recorded macros accumulate in localStorage with no way to remove
 * them).
 */
type Props = {
  api: FUniver;
  onClose: () => void;
  onRan?: (name: string, steps: number) => void;
};

export function MacrosDialog({ api, onClose, onRan }: Props) {
  const [macros, setMacros] = useState<Macro[]>(() => listMacros());
  const [busy, setBusy] = useState(false);

  const run = async (m: Macro) => {
    setBusy(true);
    try {
      const n = await runMacro(api, m.steps);
      onRan?.(m.name, n);
      onClose();
    } finally {
      setBusy(false);
    }
  };

  const remove = (name: string) => {
    setMacros(deleteMacro(name));
  };

  const assignShortcut = (name: string, letter: string) => {
    setMacros(setMacroShortcut(name, letter || null));
  };

  return (
    <Dialog title="Macros" onClose={onClose} data-testid="macros-dialog">
      {macros.length === 0 ? (
        <div className="macros-dialog__empty" data-testid="macros-dialog-empty">
          No macros yet. Use <strong>Data → Macros → Record macro</strong> to capture one.
        </div>
      ) : (
        <ul className="macros-dialog__list">
          {macros.map((m) => (
            <li
              className="macros-dialog__row"
              key={m.name}
              data-testid={`macros-dialog-row-${m.name.replace(/\s+/g, '-')}`}
            >
              <span className="macros-dialog__name">
                {m.name}
                <span className="macros-dialog__count">
                  {m.steps.length} step{m.steps.length === 1 ? '' : 's'}
                </span>
              </span>
              <label className="macros-dialog__shortcut">
                <span className="macros-dialog__shortcut-prefix">Ctrl+Shift+</span>
                <select
                  className="macros-dialog__shortcut-select"
                  data-testid={`macros-dialog-shortcut-${m.name.replace(/\s+/g, '-')}`}
                  value={m.shortcut ?? ''}
                  disabled={busy}
                  onChange={(e) => assignShortcut(m.name, e.target.value)}
                >
                  <option value="">—</option>
                  {availableMacroLetters(m.name, macros).map((l) => (
                    <option key={l} value={l}>
                      {l}
                    </option>
                  ))}
                </select>
              </label>
              <span className="macros-dialog__actions">
                <button
                  type="button"
                  className="btn-primary"
                  data-testid={`macros-dialog-run-${m.name.replace(/\s+/g, '-')}`}
                  disabled={busy}
                  onClick={() => void run(m)}
                >
                  Run
                </button>
                <button
                  type="button"
                  className="btn-secondary"
                  data-testid={`macros-dialog-delete-${m.name.replace(/\s+/g, '-')}`}
                  disabled={busy}
                  onClick={() => remove(m.name)}
                >
                  Delete
                </button>
              </span>
            </li>
          ))}
        </ul>
      )}
    </Dialog>
  );
}
