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

import { useEffect, useRef, useState } from 'react';
import { Dialog } from './Dialog';
import type { PasteSpecialMode } from './home-tab-actions';

type Props = {
  onCancel: () => void;
  onConfirm: (mode: PasteSpecialMode) => void;
};

/**
 * Excel's Paste Special dialog. Each option maps 1:1 to a Univer paste
 * hook (see `pasteSpecial` in home-tab-actions.ts). We surface the six
 * hooks Univer OSS supports — transpose / skip-blanks / arithmetic
 * operations need clipboard internals Univer doesn't expose, so they're
 * intentionally omitted rather than faked.
 */
const OPTIONS: { id: PasteSpecialMode; label: string; hint: string }[] = [
  { id: 'all', label: 'All', hint: 'Values, formulas, and formatting' },
  { id: 'formulas', label: 'Formulas', hint: 'Formulas without formatting' },
  { id: 'values', label: 'Values', hint: 'Computed values, no formulas' },
  { id: 'formats', label: 'Formats', hint: 'Formatting only, no contents' },
  { id: 'col-widths', label: 'Column widths', hint: 'Source column widths only' },
  { id: 'no-borders', label: 'All except borders', hint: 'Everything but cell borders' },
];

export function PasteSpecialDialog({ onCancel, onConfirm }: Props) {
  const [mode, setMode] = useState<PasteSpecialMode>('all');
  const activeRadioRef = useRef<HTMLInputElement | null>(null);
  useEffect(() => {
    activeRadioRef.current?.focus();
  }, []);
  const onGroupKeyDown = (e: React.KeyboardEvent<HTMLDivElement>) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      onConfirm(mode);
    } else if (e.key === 'Escape') {
      e.preventDefault();
      onCancel();
    }
  };
  return (
    <Dialog
      title="Paste Special"
      onClose={onCancel}
      data-testid="paste-special-dialog"
      footer={
        <>
          <button
            type="button"
            className="btn-secondary"
            data-testid="paste-special-cancel"
            onClick={onCancel}
          >
            Cancel
          </button>
          <button
            type="button"
            className="btn-primary"
            data-testid="paste-special-ok"
            onClick={() => onConfirm(mode)}
          >
            OK
          </button>
        </>
      }
    >
      <div
        className="cells-op"
        role="radiogroup"
        aria-label="Paste Special options"
        onKeyDown={onGroupKeyDown}
      >
        {OPTIONS.map((o) => {
          const isActive = mode === o.id;
          return (
            <label
              key={o.id}
              className={`cells-op__option${isActive ? ' cells-op__option--active' : ''}`}
              data-testid={`paste-special-${o.id}`}
            >
              <input
                ref={isActive ? activeRadioRef : undefined}
                type="radio"
                name="paste-special"
                value={o.id}
                checked={isActive}
                onChange={() => setMode(o.id)}
              />
              <span className="cells-op__label">
                <span className="cells-op__label-main">{o.label}</span>
                <span className="cells-op__label-hint">{o.hint}</span>
              </span>
            </label>
          );
        })}
      </div>
    </Dialog>
  );
}
