import { useEffect, useRef, useState, type KeyboardEvent } from 'react';
import { useWorkbook } from '../use-workbook';
import { useUI } from '../use-ui';
import { useUniverAPI } from '../use-univer';
import { useCollab } from '../collab/collab-context';
import { AvatarStack } from '../collab/AvatarStack';
import { useTheme } from '../theme';
import { Icon } from './Icon';
import { BusyPill } from './BusyPill';
import { SaveStatusPill } from './SaveStatusPill';
import { ActivityPill } from './ActivityPill';
import { MenuBar } from './MenuBar';
import { NamePill } from './NamePill';
import { AccountMenu } from '../auth/AccountMenu';
import { navigate } from '../router';

/**
 * Title bar — Google-Docs-style two-row chrome.
 *
 *   ┌──────────┬────────────────────────────┬──────────────────┐
 *   │          │  Document Name             │                  │
 *   │  Logo    │                            │  Right Actions   │
 *   │          │  File  Edit  View  Insert  │                  │
 *   └──────────┴────────────────────────────┴──────────────────┘
 *
 * Logo spans both rows on the left, actions span both rows on the
 * right, and the centre column stacks the editable filename above
 * the classic dropdown menus. This collapses the previously-separate
 * menu-bar strip into one chrome layer — fixes the "two competing
 * nav strips" perception we had before.
 *
 * MenuBar is rendered inline here (not as a sibling in App's grid)
 * so the visual reads as one block. Its keydown effects keep firing
 * because the component is still mounted, just located differently.
 *
 * Click the filename to rename inline; Enter commits, Escape reverts,
 * blur commits the current draft.
 */
export function TitleBar() {
  const { meta, renameWorkbook } = useWorkbook();
  const ui = useUI();
  const api = useUniverAPI();
  const collab = useCollab();
  const { theme, toggle: toggleTheme } = useTheme();
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
      <a
        className="titlebar__brand"
        href="/home"
        onClick={(e) => {
          // Single-click → SPA navigate to /home (My Spreadsheets) instead
          // of a full reload. Cmd/Ctrl+click + middle-click + right-click
          // all preserve the default browser behaviour (open in new tab,
          // download, etc.) so we only intercept the plain left-click.
          if (e.defaultPrevented) return;
          if (e.button !== 0) return;
          if (e.metaKey || e.ctrlKey || e.shiftKey || e.altKey) return;
          e.preventDefault();
          navigate('/home');
        }}
        aria-label="Casual Sheets — home"
        title="Casual Sheets"
      >
        <img
          // Prefix with Vite's BASE_URL so the path resolves under
          // /sheets/ on GitHub Pages and stays at / in local dev.
          src={`${import.meta.env.BASE_URL}brand.svg`}
          alt="Casual Sheets"
          className="titlebar__brand-icon"
          width={32}
          height={32}
        />
      </a>
      <div className="titlebar__center">
        <div className="titlebar__center-top">
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
        </div>
        <div className="titlebar__center-bottom">
          <MenuBar />
        </div>
      </div>
      <div className="titlebar__actions" data-testid="titlebar-actions">
        <BusyPill />
        <SaveStatusPill />
        <ActivityPill />
        <AvatarStack />
        {collab.roomId && <NamePill />}
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
        <button
          type="button"
          className="titlebar__icon-btn"
          data-testid="titlebar-theme-toggle"
          aria-label={theme === 'dark' ? 'Switch to light theme' : 'Switch to dark theme'}
          title={theme === 'dark' ? 'Switch to light theme' : 'Switch to dark theme'}
          onClick={toggleTheme}
        >
          <Icon name={theme === 'dark' ? 'light_mode' : 'dark_mode'} size="sm" />
        </button>
        <AccountMenu />
      </div>
    </header>
  );
}
