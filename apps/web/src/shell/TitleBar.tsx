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

import { useEffect, useRef, useState, type KeyboardEvent } from 'react';
import { useWorkbook } from '../use-workbook';
import { useUI } from '../use-ui';
import { useUniverAPI } from '../use-univer';
import { useCollab } from '../collab/collab-context';
import { AvatarStack } from '../collab/AvatarStack';
import { isDesktop } from '../desk-bridge-bootstrap';
import { useTheme } from '../theme';
// Design-system primitives (Phase 4) — the title-bar action cluster now uses
// the shared Button / IconButton / Badge instead of hand-rolled buttons, so
// the accent gradient, hover spring, and pill tones match the system.
import { Badge, Button, IconButton } from '@schnsrw/design-system';
import { BusyPill } from './BusyPill';
import { SaveStatusPill } from './SaveStatusPill';
import { ActivityPill } from './ActivityPill';
import { MenuBar } from './MenuBar';
import { NamePill } from './NamePill';
import { AccountMenu } from '../auth/AccountMenu';
import { navigate } from '../router';
// Suite brand mark from the shared design system (Phase 4) — one logo across
// sheet/doc/slides/drive. Vite resolves the SVG import to a URL.
import sheetsMark from '@schnsrw/design-system/assets/casual-sheets-mark.svg';

type DesktopProfile = { name: string; avatar_hue: number; avatar_path: string | null };

/**
 * Fetch the local-user profile from the desktop shell (via the bridge's
 * `getProfile`). Only runs in the desktop build; web stays null so nothing
 * renders. The shell owns the profile (name + avatar hue) in profile.json.
 */
function useDesktopProfile(): { profile: DesktopProfile | null } {
  const [profile, setProfile] = useState<DesktopProfile | null>(null);
  useEffect(() => {
    if (!isDesktop()) return;
    const bridge = (window as unknown as { __deskApp__?: { getProfile?: () => Promise<unknown> } })
      .__deskApp__;
    if (typeof bridge?.getProfile !== 'function') return;
    let cancelled = false;
    bridge
      .getProfile()
      .then((p) => {
        if (!cancelled && p) setProfile(p as DesktopProfile);
      })
      .catch(() => undefined);
    return () => {
      cancelled = true;
    };
  }, []);
  return { profile };
}

function DesktopProfileChip({ profile }: { profile: DesktopProfile }) {
  const initials =
    profile.name
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
  const desktop = useDesktopProfile();
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
          // Shared design-system suite mark (Phase 4).
          src={sheetsMark}
          alt="Casual Sheets"
          className="titlebar__brand-icon"
          width={28}
          height={36}
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
              // The filename text IS the button's accessible name (announced by
              // screen readers and used by getByRole name lookups). Don't set
              // aria-label here — it would override the filename with a generic
              // string. The rename affordance is conveyed via the title tooltip.
              title={`${filename} — click to rename`}
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
        {/* App status — transient document state. */}
        <BusyPill />
        <SaveStatusPill />
        <ActivityPill />
        {/* Collaboration cluster — presence + who-you-are + share, grouped
            together (Docs/Excel-online style) so the co-edit affordances read
            as one unit, distinct from the status pills and account controls.
            Hidden in the desktop build: it's a single-user, local-file app with
            no co-editing, so Share / avatars / room pill don't apply. */}
        {!isDesktop() && (
          <>
            <span className="titlebar__sep" aria-hidden="true" />
            <div className="titlebar__collab" data-testid="titlebar-collab">
              <AvatarStack />
              {collab.roomId && <NamePill />}
              {collab.roomId ? (
                <Badge
                  tone="accent"
                  icon="group"
                  data-testid="titlebar-roompill"
                  title={`Joined room ${collab.roomId}`}
                >
                  In room
                </Badge>
              ) : (
                <Button
                  variant="primary"
                  size="sm"
                  icon="group_add"
                  data-testid="titlebar-share"
                  onClick={() => ui.openShareRoom()}
                >
                  Share
                </Button>
              )}
            </div>
          </>
        )}
        {/* Desktop build: a "Home" affordance that brings the Casual Office
            launcher window forward (the launcher IS the home screen). Desktop
            only — web has its own /home brand link above. */}
        {isDesktop() && (
          <IconButton
            size="md"
            icon="home"
            label="Back to Casual Office"
            data-testid="titlebar-back-to-launcher"
            onClick={() => {
              try {
                // Top-level desktop mode: __TAURI__.core.invoke is direct.
                (
                  window as unknown as {
                    __TAURI__?: { core?: { invoke?: (cmd: string) => Promise<unknown> } };
                  }
                ).__TAURI__?.core?.invoke?.('focus_launcher_window')?.catch?.(() => undefined);
              } catch {
                /* best-effort — never break the title bar */
              }
            }}
          />
        )}
        {/* Desktop build: a local-user chip stands in for the collab cluster —
            single-user, so it just shows who you're signed in as (from the
            shell's profile.json via the bridge). */}
        {isDesktop() && desktop.profile && <DesktopProfileChip profile={desktop.profile} />}
        <span className="titlebar__sep" aria-hidden="true" />
        <IconButton
          size="md"
          icon={theme === 'dark' ? 'light_mode' : 'dark_mode'}
          label={theme === 'dark' ? 'Switch to light theme' : 'Switch to dark theme'}
          data-testid="titlebar-theme-toggle"
          onClick={toggleTheme}
        />
        <AccountMenu />
      </div>
    </header>
  );
}
