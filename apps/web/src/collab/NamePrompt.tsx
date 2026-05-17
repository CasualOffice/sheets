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
            You can change it later from the share menu.
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
