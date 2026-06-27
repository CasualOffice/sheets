import { useContext, useEffect, useRef, useState } from 'react';
import { confirmModal } from '../shell/modals';
import { Icon } from '../shell/Icon';
import { useAuth, useCurrentUser } from './auth-context';
import { logout } from './api';
import { SettingsModal } from './SettingsModal';
import { avatarUrl } from './profile-api';
import { WorkbookContext } from '../workbook-context';

/**
 * Title-bar account menu — only renders when a personal-mode
 * session is live (Mode 3 / authenticated). Avatar circle with the
 * first letter of the username; click opens a small popover with
 * the username + Sign out. Phase C-follow-up: "Change password"
 * and "Delete account" land here once the modal is built.
 *
 * Hidden entirely on Mode 1 / Mode 2 / unauthenticated states so
 * the existing chrome layout is unchanged outside of the
 * standalone deploy.
 */

export function AccountMenu() {
  const user = useCurrentUser();
  const { setUnauthenticated, state } = useAuth();
  const [open, setOpen] = useState(false);
  const [settingsOpen, setSettingsOpen] = useState(false);
  // Bumps when the modal saves a new avatar so the AccountMenu's
  // <img> tag re-fetches without sleeping on the HTTP cache.
  const [avatarBust, setAvatarBust] = useState(0);
  // The <img> renders eagerly; onError swaps to the initial. This
  // avoids the upfront HEAD probe on every chrome render.
  const [avatarOk, setAvatarOk] = useState(true);
  const ref = useRef<HTMLDivElement | null>(null);
  void state;

  useEffect(() => {
    if (!open) return;
    const onDown = (e: MouseEvent) => {
      if (ref.current && !ref.current.contains(e.target as Node)) setOpen(false);
    };
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setOpen(false);
    };
    document.addEventListener('mousedown', onDown);
    document.addEventListener('keydown', onKey);
    return () => {
      document.removeEventListener('mousedown', onDown);
      document.removeEventListener('keydown', onKey);
    };
  }, [open]);

  // Reset the avatar-ok flag when the user changes or when the modal
  // bumps the cache-buster, so a freshly-uploaded image is re-tried
  // even if the previous URL hit the onError fallback. Must run as a
  // hook (not gated by `!user`) to satisfy the Rules of Hooks.
  useEffect(() => {
    setAvatarOk(true);
  }, [user?.id, avatarBust]);

  // Read the workbook context directly (not via useWorkbook) because
  // AccountMenu is rendered on `/home` too — where WorkbookContext is
  // still mounted, but a future code path might not provide one. Plain
  // `useContext` returns null instead of throwing, which is the right
  // shape for an "is there anything to lose?" probe.
  const wbCtx = useContext(WorkbookContext);

  if (!user) return null;

  const initial = user.username.charAt(0).toUpperCase();
  const onLogout = async () => {
    // UX_AUDIT.md §2.14 — confirm before discarding unsaved edits. The
    // dirty flag is reset on every successful save, so this only fires
    // when the user has typed since the last save (or since opening a
    // fresh draft).
    if (wbCtx?.meta.hasUserEdited) {
      const ok = await confirmModal({
        title: 'Sign out?',
        body: 'You have unsaved changes. Sign out anyway? Your edits will be lost.',
        confirmLabel: 'Sign out',
        danger: true,
      });
      if (!ok) return;
    }
    setOpen(false);
    // Clear local auth state FIRST so the gate swaps to the login view
    // immediately, independent of how long the server round-trip takes. Awaiting
    // logout() before this made the login view's appearance hostage to network
    // latency — the source of an intermittent personal-mode e2e failure where
    // `auth-login` didn't show within the 10s wait on a slow CI runner. The
    // server session clear is best-effort; on failure the cookie expires or the
    // next `/auth/status` probe reconciles.
    setUnauthenticated();
    try {
      await logout();
    } catch {
      /* local state already cleared; nothing actionable here */
    }
  };

  return (
    <>
      <div className="account-menu" ref={ref} data-testid="account-menu">
        <button
          type="button"
          className="account-menu__trigger"
          aria-haspopup="menu"
          aria-expanded={open}
          title={`Signed in as ${user.username}`}
          onClick={() => setOpen((v) => !v)}
          data-testid="account-menu-trigger"
        >
          {avatarOk ? (
            <img
              src={avatarUrl(user.id, avatarBust || undefined)}
              alt=""
              className="account-menu__avatar-img"
              onError={() => setAvatarOk(false)}
            />
          ) : (
            <span aria-hidden>{initial}</span>
          )}
        </button>
        {open && (
          <div role="menu" className="account-menu__popover">
            <div className="account-menu__header">
              <div className="account-menu__name">{user.username}</div>
              <div className="account-menu__role">{user.isAdmin ? 'Admin' : 'User'}</div>
            </div>
            <button
              type="button"
              role="menuitem"
              className="account-menu__item"
              onClick={() => {
                setOpen(false);
                setSettingsOpen(true);
              }}
              data-testid="account-menu-settings"
            >
              <Icon name="settings" size="sm" />
              <span>Settings</span>
            </button>
            {user.isAdmin && (
              // Admin panel lives in a separate React entry point
              // (apps/web/src/admin/AdminApp.tsx) selected at boot
              // from `window.location.pathname.startsWith('/admin')`,
              // so it's not router-navigable from this tree — use a
              // plain anchor for a full reload. UX_AUDIT.md §2.7.
              <a
                role="menuitem"
                className="account-menu__item"
                href="/admin"
                onClick={() => setOpen(false)}
                data-testid="account-menu-admin"
              >
                <Icon name="settings_applications" size="sm" />
                <span>Admin panel</span>
              </a>
            )}
            <button
              type="button"
              role="menuitem"
              className="account-menu__item"
              onClick={() => void onLogout()}
              data-testid="account-menu-signout"
            >
              <Icon name="logout" size="sm" />
              <span>Sign out</span>
            </button>
          </div>
        )}
      </div>
      {settingsOpen && (
        <SettingsModal
          onClose={() => {
            setSettingsOpen(false);
            // Bump the cache buster so a fresh avatar shows up immediately
            // without waiting for the 2 min cache-control header.
            setAvatarBust(Date.now());
          }}
        />
      )}
    </>
  );
}
