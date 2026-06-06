import { useEffect, useRef, useState } from 'react';
import { Icon } from '../shell/Icon';
import { useAuth, useCurrentUser } from './auth-context';
import { logout } from './api';
import { SettingsModal } from './SettingsModal';
import { avatarUrl } from './profile-api';

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
  const ref = useRef<HTMLDivElement | null>(null);
  // Detect whether the user has uploaded an avatar by hitting the
  // image and falling back to the initial on error — saves an
  // upfront /auth/profile probe on every chrome render.
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

  if (!user) return null;

  const initial = user.username.charAt(0).toUpperCase();
  const onLogout = async () => {
    setOpen(false);
    await logout();
    setUnauthenticated();
  };

  // The <img> renders eagerly; onError swaps to the initial. This
  // avoids the upfront HEAD probe on every chrome render.
  const [avatarOk, setAvatarOk] = useState(true);
  useEffect(() => {
    setAvatarOk(true);
  }, [user.id, avatarBust]);

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
