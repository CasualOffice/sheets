import { useEffect, useRef, useState } from 'react';
import { Icon } from '../shell/Icon';
import { useAuth } from './auth-context';
import { changePassword, deleteAccount, logout } from './api';
import {
  avatarUrl,
  deleteAvatar,
  fetchProfile,
  patchProfile,
  uploadAvatar,
  type UserProfile,
} from './profile-api';

/**
 * Tabbed settings dialog reached from the AccountMenu. Three tabs:
 *
 *   Profile      — display name, email, timezone, avatar
 *   Security     — current password → new password, delete account
 *   Preferences  — theme, language (round-tripped via preferences JSON)
 *
 * Pure modal — escape + backdrop click + the X close. Saves are
 * per-tab; each tab handles its own busy/error state so a slow
 * avatar upload doesn't lock out a separate password change. The
 * AccountMenu rerenders when this modal saves (via auth.refresh)
 * so the avatar circle picks up changes.
 */

export function SettingsModal({ onClose }: { onClose: () => void }) {
  const { refresh, setUnauthenticated } = useAuth();
  const [tab, setTab] = useState<'profile' | 'security' | 'preferences'>('profile');
  const [profile, setProfile] = useState<UserProfile | null>(null);
  const [userId, setUserId] = useState<number | null>(null);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    void fetchProfile().then(
      ({ user, profile }) => {
        if (cancelled) return;
        setUserId(user.id);
        setProfile(profile);
      },
      (err: unknown) => {
        if (!cancelled) setError(err instanceof Error ? err.message : String(err));
      },
    );
    return () => {
      cancelled = true;
    };
  }, []);

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose();
    };
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  }, [onClose]);

  return (
    <div
      className="settings-modal__backdrop"
      data-testid="settings-modal"
      onClick={(e) => {
        if (e.target === e.currentTarget) onClose();
      }}
    >
      <div className="settings-modal" role="dialog" aria-label="Settings" aria-modal="true">
        <header className="settings-modal__head">
          <h1>Settings</h1>
          <button
            type="button"
            className="settings-modal__close"
            aria-label="Close settings"
            onClick={onClose}
            data-testid="settings-close"
          >
            <Icon name="close" />
          </button>
        </header>

        <div className="settings-modal__body">
          <nav className="settings-modal__tabs" aria-label="Settings tabs">
            {(
              [
                ['profile', 'Profile', 'person'],
                ['security', 'Security', 'lock'],
                ['preferences', 'Preferences', 'tune'],
              ] as const
            ).map(([id, label, icon]) => (
              <button
                key={id}
                type="button"
                role="tab"
                aria-selected={tab === id}
                className={`settings-modal__tab${tab === id ? ' settings-modal__tab--on' : ''}`}
                onClick={() => setTab(id)}
                data-testid={`settings-tab-${id}`}
              >
                <Icon name={icon} size="sm" />
                <span>{label}</span>
              </button>
            ))}
          </nav>

          <section className="settings-modal__panel">
            {error && (
              <div className="auth-card__error" role="alert">
                {error}
              </div>
            )}
            {profile && userId !== null && tab === 'profile' && (
              <ProfileTab
                userId={userId}
                profile={profile}
                onSaved={(p) => {
                  setProfile(p);
                  void refresh();
                }}
              />
            )}
            {profile && tab === 'security' && (
              <SecurityTab
                onLoggedOut={() => {
                  setUnauthenticated();
                  onClose();
                }}
                onDeleted={async () => {
                  await logout();
                  setUnauthenticated();
                  onClose();
                }}
              />
            )}
            {profile && tab === 'preferences' && (
              <PreferencesTab
                profile={profile}
                onSaved={(p) => {
                  setProfile(p);
                  void refresh();
                }}
              />
            )}
          </section>
        </div>
      </div>
    </div>
  );
}

// ── Profile tab ────────────────────────────────────────────────────────

function ProfileTab({
  userId,
  profile,
  onSaved,
}: {
  userId: number;
  profile: UserProfile;
  onSaved: (next: UserProfile) => void;
}) {
  const [displayName, setDisplayName] = useState(profile.displayName ?? '');
  const [email, setEmail] = useState(profile.email ?? '');
  const [timezone, setTimezone] = useState(profile.timezone);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [avatarBust, setAvatarBust] = useState(0);
  const fileRef = useRef<HTMLInputElement>(null);
  const hasAvatar = profile.hasAvatar || avatarBust > 0;

  const tzList = supportedTimeZones();

  const onSave = async () => {
    setBusy(true);
    setError(null);
    const result = await patchProfile({
      displayName: displayName.trim() || null,
      email: email.trim() || null,
      timezone,
    });
    setBusy(false);
    if (!result.ok || !result.profile) {
      setError(humanise(result.reason ?? 'unknown'));
      return;
    }
    onSaved(result.profile);
  };

  const onAvatar = async (file: File) => {
    if (file.size > 256 * 1024) {
      setError('Avatar must be 256 KB or smaller.');
      return;
    }
    setBusy(true);
    const r = await uploadAvatar(file);
    setBusy(false);
    if (!r.ok) {
      setError(humanise(r.reason ?? 'unknown'));
      return;
    }
    setAvatarBust(Date.now());
    onSaved({ ...profile, hasAvatar: true });
  };

  return (
    <div className="settings-tab" data-testid="settings-tab-profile-content">
      <div className="settings-avatar">
        <div className="settings-avatar__preview">
          {hasAvatar ? (
            <img
              src={avatarUrl(userId, avatarBust || undefined)}
              alt="Avatar"
              data-testid="settings-avatar-img"
            />
          ) : (
            <span className="settings-avatar__initial" aria-hidden>
              {(displayName || profile.email || 'U').charAt(0).toUpperCase()}
            </span>
          )}
        </div>
        <div className="settings-avatar__actions">
          <button
            type="button"
            className="auth-btn"
            onClick={() => fileRef.current?.click()}
            data-testid="settings-avatar-upload"
          >
            <Icon name="upload" size="sm" />
            <span>Upload</span>
          </button>
          {hasAvatar && (
            <button
              type="button"
              className="auth-btn"
              data-testid="settings-avatar-remove"
              onClick={async () => {
                await deleteAvatar();
                setAvatarBust(Date.now());
                onSaved({ ...profile, hasAvatar: false });
              }}
            >
              <Icon name="delete" size="sm" />
              <span>Remove</span>
            </button>
          )}
          <input
            ref={fileRef}
            type="file"
            accept="image/png,image/jpeg,image/webp,image/gif"
            hidden
            data-testid="settings-avatar-input"
            onChange={(e) => {
              const f = e.target.files?.[0];
              if (f) void onAvatar(f);
              e.target.value = '';
            }}
          />
        </div>
      </div>

      <label className="auth-card__field">
        <span>Display name</span>
        <input
          type="text"
          maxLength={80}
          value={displayName}
          onChange={(e) => setDisplayName(e.target.value)}
          data-testid="settings-displayname"
          placeholder="How others see you"
        />
      </label>

      <label className="auth-card__field">
        <span>Email (optional)</span>
        <input
          type="email"
          autoComplete="email"
          value={email}
          onChange={(e) => setEmail(e.target.value)}
          data-testid="settings-email"
          placeholder="alice@example.com"
        />
      </label>

      <label className="auth-card__field">
        <span>Timezone</span>
        <select
          value={timezone}
          onChange={(e) => setTimezone(e.target.value)}
          data-testid="settings-timezone"
        >
          {tzList.map((tz) => (
            <option key={tz} value={tz}>
              {tz}
            </option>
          ))}
        </select>
      </label>

      {error && (
        <div className="auth-card__error" role="alert" data-testid="settings-error">
          {error}
        </div>
      )}

      <div className="settings-tab__actions">
        <button
          type="button"
          className="auth-btn auth-btn--primary"
          onClick={() => void onSave()}
          disabled={busy}
          data-testid="settings-save-profile"
        >
          {busy ? 'Saving…' : 'Save changes'}
        </button>
      </div>
    </div>
  );
}

// ── Security tab ───────────────────────────────────────────────────────

function SecurityTab({
  onLoggedOut,
  onDeleted,
}: {
  onLoggedOut: () => void;
  onDeleted: () => void;
}) {
  const [current, setCurrent] = useState('');
  const [next, setNext] = useState('');
  const [confirmNext, setConfirmNext] = useState('');
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [confirmDelete, setConfirmDelete] = useState(false);

  const onSave = async () => {
    setError(null);
    if (next.length < 8) {
      setError('New password must be at least 8 characters.');
      return;
    }
    if (next !== confirmNext) {
      setError('Passwords don’t match.');
      return;
    }
    setBusy(true);
    const r = await changePassword(current, next);
    setBusy(false);
    if (!r.ok) {
      setError(humanise(r.reason ?? 'unknown'));
      return;
    }
    setCurrent('');
    setNext('');
    setConfirmNext('');
    onLoggedOut();
  };

  const onDelete = async () => {
    setBusy(true);
    const r = await deleteAccount();
    setBusy(false);
    if (!r.ok) {
      setError(humanise(r.reason ?? 'unknown'));
      return;
    }
    onDeleted();
  };

  return (
    <div className="settings-tab" data-testid="settings-tab-security-content">
      <h2 className="settings-tab__heading">Change password</h2>
      <label className="auth-card__field">
        <span>Current password</span>
        <input
          type="password"
          autoComplete="current-password"
          value={current}
          onChange={(e) => setCurrent(e.target.value)}
          data-testid="settings-current-password"
        />
      </label>
      <label className="auth-card__field">
        <span>New password</span>
        <input
          type="password"
          autoComplete="new-password"
          value={next}
          minLength={8}
          onChange={(e) => setNext(e.target.value)}
          data-testid="settings-new-password"
        />
      </label>
      <label className="auth-card__field">
        <span>Confirm new password</span>
        <input
          type="password"
          autoComplete="new-password"
          value={confirmNext}
          minLength={8}
          onChange={(e) => setConfirmNext(e.target.value)}
          data-testid="settings-confirm-password"
        />
      </label>

      {error && (
        <div className="auth-card__error" role="alert" data-testid="settings-error">
          {error}
        </div>
      )}

      <div className="settings-tab__actions">
        <button
          type="button"
          className="auth-btn auth-btn--primary"
          onClick={() => void onSave()}
          disabled={busy}
          data-testid="settings-save-password"
        >
          {busy ? 'Updating…' : 'Update password'}
        </button>
      </div>

      <h2 className="settings-tab__heading settings-tab__heading--danger">Danger zone</h2>
      {!confirmDelete ? (
        <div className="settings-tab__danger">
          <p>Delete your account and every file you own. This can’t be undone.</p>
          <button
            type="button"
            className="auth-btn settings-tab__danger-btn"
            onClick={() => setConfirmDelete(true)}
            data-testid="settings-delete-trigger"
          >
            Delete my account
          </button>
        </div>
      ) : (
        <div className="settings-tab__danger">
          <p>
            Are you sure? This wipes your files, sessions, and credentials. The server keeps no
            backup.
          </p>
          <div className="settings-tab__actions">
            <button
              type="button"
              className="auth-btn"
              onClick={() => setConfirmDelete(false)}
              data-testid="settings-delete-cancel"
            >
              Cancel
            </button>
            <button
              type="button"
              className="auth-btn settings-tab__danger-btn"
              onClick={() => void onDelete()}
              disabled={busy}
              data-testid="settings-delete-confirm"
            >
              Yes, delete my account
            </button>
          </div>
        </div>
      )}
    </div>
  );
}

// ── Preferences tab ───────────────────────────────────────────────────

type Prefs = {
  theme?: 'light' | 'dark' | 'system';
  locale?: string;
  dateFormat?: string;
};

function PreferencesTab({
  profile,
  onSaved,
}: {
  profile: UserProfile;
  onSaved: (next: UserProfile) => void;
}) {
  const prefs = (profile.preferences ?? {}) as Prefs;
  const [theme, setTheme] = useState<Prefs['theme']>(prefs.theme ?? 'system');
  const [locale, setLocale] = useState(prefs.locale ?? 'en');
  const [dateFormat, setDateFormat] = useState(prefs.dateFormat ?? 'yyyy-mm-dd');
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const onSave = async () => {
    setBusy(true);
    setError(null);
    const result = await patchProfile({
      preferences: { ...prefs, theme, locale, dateFormat },
    });
    setBusy(false);
    if (!result.ok || !result.profile) {
      setError(humanise(result.reason ?? 'unknown'));
      return;
    }
    onSaved(result.profile);
  };

  return (
    <div className="settings-tab" data-testid="settings-tab-preferences-content">
      <label className="auth-card__field">
        <span>Theme</span>
        <select
          value={theme}
          onChange={(e) => setTheme(e.target.value as Prefs['theme'])}
          data-testid="settings-theme"
        >
          <option value="system">Match system</option>
          <option value="light">Light</option>
          <option value="dark">Dark</option>
        </select>
      </label>

      <label className="auth-card__field">
        <span>Language</span>
        <select
          value={locale}
          onChange={(e) => setLocale(e.target.value)}
          data-testid="settings-locale"
        >
          <option value="en">English</option>
          <option value="en-GB">English (UK)</option>
          <option value="es">Español</option>
          <option value="de">Deutsch</option>
          <option value="fr">Français</option>
          <option value="ja">日本語</option>
          <option value="zh">中文</option>
        </select>
      </label>

      <label className="auth-card__field">
        <span>Date format</span>
        <select
          value={dateFormat}
          onChange={(e) => setDateFormat(e.target.value)}
          data-testid="settings-date-format"
        >
          <option value="yyyy-mm-dd">2026-06-06 (ISO)</option>
          <option value="mm/dd/yyyy">06/06/2026 (US)</option>
          <option value="dd/mm/yyyy">06/06/2026 (UK / Europe)</option>
          <option value="d mmm yyyy">6 Jun 2026</option>
        </select>
      </label>

      {error && (
        <div className="auth-card__error" role="alert" data-testid="settings-error">
          {error}
        </div>
      )}

      <div className="settings-tab__actions">
        <button
          type="button"
          className="auth-btn auth-btn--primary"
          onClick={() => void onSave()}
          disabled={busy}
          data-testid="settings-save-preferences"
        >
          {busy ? 'Saving…' : 'Save preferences'}
        </button>
      </div>
    </div>
  );
}

// ── Helpers ───────────────────────────────────────────────────────────

function humanise(reason: string): string {
  switch (reason) {
    case 'conflict-or-invalid':
      return 'Couldn’t save — check the fields and try again.';
    case 'rejected':
      return 'Current password is wrong, or the new one is too short.';
    case 'last-admin':
      return 'You’re the last admin — promote another user before deleting your account.';
    case 'avatar-too-large':
      return 'Avatar must be 256 KB or smaller.';
    case 'unsupported-mime':
      return 'Use a PNG, JPEG, WebP, or GIF image.';
    default:
      return `Couldn’t save (${reason}).`;
  }
}

function supportedTimeZones(): string[] {
  // Prefer the platform's list when available; fall back to a small
  // hand-picked set that covers most users without a 600-entry select.
  const intlAny = Intl as unknown as { supportedValuesOf?: (key: string) => string[] };
  if (typeof intlAny.supportedValuesOf === 'function') {
    try {
      return intlAny.supportedValuesOf('timeZone');
    } catch {
      /* fall through */
    }
  }
  return [
    'UTC',
    'Europe/London',
    'Europe/Berlin',
    'Europe/Paris',
    'Europe/Moscow',
    'Africa/Cairo',
    'Asia/Dubai',
    'Asia/Kolkata',
    'Asia/Shanghai',
    'Asia/Tokyo',
    'Australia/Sydney',
    'America/New_York',
    'America/Chicago',
    'America/Denver',
    'America/Los_Angeles',
    'America/Sao_Paulo',
  ];
}
