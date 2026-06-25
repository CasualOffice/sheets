import { useState } from 'react';
import type { FUniver } from '@univerjs/core/facade';
import { Dialog } from './Dialog';
import {
  availableMacroLetters,
  deleteMacro,
  listMacros,
  renameMacro,
  runMacro,
  setMacroShortcut,
  type Macro,
} from '../sheets/macros';

/**
 * Inline-editable macro name. Commits on Enter / blur, reverts on Escape or
 * when the rename is rejected (blank / duplicate). Local draft so typing
 * doesn't churn the parent until commit.
 */
function MacroNameInput({
  name,
  disabled,
  onRename,
}: {
  name: string;
  disabled: boolean;
  onRename: (oldName: string, newName: string) => boolean;
}) {
  const [draft, setDraft] = useState(name);
  const commit = () => {
    if (draft.trim() === name) {
      setDraft(name); // normalize any whitespace-only change back
      return;
    }
    if (!onRename(name, draft)) setDraft(name); // rejected → revert
  };
  return (
    <input
      className="macros-dialog__name-input"
      data-testid={`macros-dialog-name-${name.replace(/\s+/g, '-')}`}
      value={draft}
      disabled={disabled}
      aria-label={`Macro name: ${name}`}
      spellCheck={false}
      onChange={(e) => setDraft(e.target.value)}
      onBlur={commit}
      onKeyDown={(e) => {
        if (e.key === 'Enter') {
          e.preventDefault();
          e.currentTarget.blur();
        } else if (e.key === 'Escape') {
          setDraft(name);
          e.currentTarget.blur();
        }
      }}
    />
  );
}

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

  const rename = (oldName: string, newName: string): boolean => {
    const next = renameMacro(oldName, newName);
    setMacros(next);
    return next.some((m) => m.name === newName.trim());
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
                <MacroNameInput name={m.name} disabled={busy} onRename={rename} />
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
