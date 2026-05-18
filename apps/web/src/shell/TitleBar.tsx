import { useEffect, useRef, useState, type KeyboardEvent } from 'react';
import { useWorkbook } from '../use-workbook';
import { useUI } from '../use-ui';
import { useUniverAPI } from '../use-univer';
import { useCollab } from '../collab/collab-context';
import { AvatarStack } from '../collab/AvatarStack';
import { Icon } from './Icon';
import { BusyPill } from './BusyPill';

/**
 * Title bar — brand on the left, editable filename in the middle. Click the
 * filename to rename inline (Google-Sheets pattern); Enter commits, Escape
 * reverts, blur commits the current draft.
 */
export function TitleBar() {
  const { meta, renameWorkbook } = useWorkbook();
  const ui = useUI();
  const api = useUniverAPI();
  const collab = useCollab();
  const filename = meta.name || 'Untitled';

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
    if (next && next !== meta.name) {
      // In-place rename via Univer + meta context — no unit swap, no
      // snapshot clone. App owns the meta update; we mirror into Univer
      // here since App is outside the UniverProvider.
      try {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        (api?.getActiveWorkbook?.() as any)?.setName?.(next);
      } catch {
        /* facade missing setName — meta still updates below */
      }
      renameWorkbook(next);
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
          // Prefix with Vite's BASE_URL so the path resolves under
          // /sheets/ on GitHub Pages and stays at / in local dev.
          src={`${import.meta.env.BASE_URL}brand.svg`}
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

      <div className="titlebar__actions" data-testid="titlebar-actions">
        <BusyPill />
        <AvatarStack />
        {collab.roomId ? (
          <span
            className="titlebar__roompill"
            data-testid="titlebar-roompill"
            title={`Joined room ${collab.roomId}`}
          >
            <Icon name="group" size="sm" />
            <span>In room</span>
          </span>
        ) : (
          <button
            type="button"
            className="titlebar__share btn-primary"
            data-testid="titlebar-share"
            onClick={() => ui.openShareRoom()}
          >
            <Icon name="group_add" size="sm" />
            <span>Share</span>
          </button>
        )}
      </div>
    </header>
  );
}
