import { useEffect, useRef, useState, type KeyboardEvent } from 'react';
import { useWorkbook } from '../use-workbook';

/**
 * Title bar — brand on the left, editable filename in the middle. Click the
 * filename to rename inline (Google-Sheets pattern); Enter commits, Escape
 * reverts, blur commits the current draft.
 */
export function TitleBar() {
  const { snapshot, replaceWorkbook } = useWorkbook();
  const filename = snapshot.name || 'Untitled';

  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState(filename);
  const inputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    if (editing) {
      inputRef.current?.focus();
      inputRef.current?.select();
    }
  }, [editing]);

  // Keep draft synced if filename changes externally (e.g. Open xlsx).
  useEffect(() => {
    if (!editing) setDraft(filename);
  }, [filename, editing]);

  const commit = () => {
    const next = draft.trim();
    if (next && next !== snapshot.name) {
      replaceWorkbook({ ...snapshot, name: next });
    }
    setEditing(false);
  };

  const revert = () => {
    setDraft(filename);
    setEditing(false);
  };

  const onKeyDown = (e: KeyboardEvent<HTMLInputElement>) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      commit();
    } else if (e.key === 'Escape') {
      e.preventDefault();
      revert();
    }
  };

  return (
    <header className="titlebar" data-testid="titlebar" role="banner">
      <a className="titlebar__brand" href="/" aria-label="Casual Sheets — home">
        <img
          src="/brand.svg"
          alt=""
          className="titlebar__brand-icon"
          width={28}
          height={28}
        />
        <span className="titlebar__brand-name">Casual Sheets</span>
      </a>
      <span className="titlebar__divider" aria-hidden="true" />
      {editing ? (
        <input
          ref={inputRef}
          type="text"
          className="titlebar__filename-input"
          data-testid="titlebar-filename-input"
          value={draft}
          maxLength={120}
          aria-label="Rename file"
          onChange={(e) => setDraft(e.target.value)}
          onBlur={commit}
          onKeyDown={onKeyDown}
        />
      ) : (
        <button
          type="button"
          className="titlebar__filename"
          data-testid="titlebar-filename"
          title="Rename"
          onClick={() => setEditing(true)}
        >
          {filename}
        </button>
      )}
      <span className="titlebar__spacer" />
    </header>
  );
}
