import { useEffect, useRef, useState, type KeyboardEvent } from 'react';
import { useWorkbook } from '../use-workbook';
import { useUI } from '../use-ui';
import { useUniverAPI } from '../use-univer';
import { useCollab } from '../collab/collab-context';
import { AvatarStack } from '../collab/AvatarStack';
import { Icon } from './Icon';
import { BusyPill } from './BusyPill';

/**
 * True when the editor is mounted inside the Casual Office Tauri shell
 * (window.__deskApp__ is wired by the bootstrap from ?desk=1). In that
 * case the workbook is local + single-user — collab UI (Share button,
 * AvatarStack, "In room" pill) is dead weight that confuses the user,
 * so we hide it and show a local-user profile chip instead.
 */
function useDesktopProfile() {
  const [profile, setProfile] = useState<{
    name: string;
    avatar_hue: number;
    avatar_path: string | null;
  } | null>(null);
  const isDesktop =
    typeof window !== 'undefined' &&
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (window as any).__deskApp__?.isDesktop === true;
  useEffect(() => {
    if (!isDesktop) return;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const bridge = (window as any).__deskApp__;
    if (typeof bridge.getProfile !== 'function') return;
    let cancelled = false;
    bridge
      .getProfile()
      .then((p: typeof profile) => { if (!cancelled) setProfile(p); })
      .catch(() => undefined);
    return () => { cancelled = true; };
  }, [isDesktop]);
  return { isDesktop, profile };
}

function DesktopProfileChip({
  profile,
}: {
  profile: { name: string; avatar_hue: number; avatar_path: string | null };
}) {
  const initials = profile.name
    .trim()
    .split(/\s+/)
    .slice(0, 2)
    .map((p) => p[0]?.toUpperCase() ?? '')
    .join('') || '?';
  const first = profile.name.split(/\s+/)[0] ?? '';
  return (
    <div
      className="titlebar__profile-chip"
      data-testid="titlebar-profile-chip"
      title={profile.name}
    >
      <span
        className="titlebar__profile-avatar"
        style={{ backgroundColor: `hsl(${profile.avatar_hue}, 55%, 50%)` }}
        aria-hidden="true"
      >
        {initials}
      </span>
      <span className="titlebar__profile-name">{first}</span>
    </div>
  );
}

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
  const desktop = useDesktopProfile();
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
        {desktop.isDesktop ? (
          desktop.profile ? <DesktopProfileChip profile={desktop.profile} /> : null
        ) : (
          <>
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
          </>
        )}
      </div>
    </header>
  );
}
