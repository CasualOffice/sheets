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

/**
 * One-time prompt on first room join, asking the user what name to show
 * peers. Stored in localStorage so subsequent rooms skip the prompt.
 * Render-blocking (dialog), but cancel/Esc accepts the suggested anon
 * name — never traps the user out of a room.
 */
export function NamePrompt({
  suggestion,
  onSubmit,
  onCancel,
}: {
  suggestion: string;
  onSubmit: (name: string) => void;
  onCancel: () => void;
}) {
  const [value, setValue] = useState(suggestion);
  const submit = () => {
    const v = value.trim();
    onSubmit(v.length > 0 ? v : suggestion);
  };
  return (
    <div className="dialog-backdrop" data-testid="name-prompt-backdrop">
      <div
        className="dialog dialog--narrow"
        role="dialog"
        aria-modal="true"
        aria-labelledby="name-prompt-title"
        data-testid="name-prompt-dialog"
      >
        <div className="dialog__header">
          <h2 className="dialog__title" id="name-prompt-title">
            What should we call you?
          </h2>
        </div>
        <div className="dialog__body">
          <p style={{ margin: '0 0 12px', fontSize: 14, lineHeight: 1.5 }}>
            Other people in the room will see this name next to your cursor.
            You can change it later by clicking the name pill in the title bar.
          </p>
          <input
            autoFocus
            type="text"
            maxLength={32}
            className="page-setup__select"
            data-testid="name-prompt-input"
            value={value}
            onChange={(e) => setValue(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Enter') submit();
              if (e.key === 'Escape') onCancel();
            }}
          />
        </div>
        <div className="dialog__footer">
          <button
            type="button"
            className="btn-secondary"
            data-testid="name-prompt-skip"
            onClick={onCancel}
          >
            Use “{suggestion}”
          </button>
          <button
            type="button"
            className="btn-primary"
            data-testid="name-prompt-submit"
            onClick={submit}
          >
            Continue
          </button>
        </div>
      </div>
    </div>
  );
}
