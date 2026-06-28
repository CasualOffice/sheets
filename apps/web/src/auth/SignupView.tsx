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
import { useAuth } from './auth-context';
import { signup } from './api';

/**
 * First-launch signup. The first account becomes admin (server-side
 * decision); the copy reinforces that — "this is your admin
 * account, choose a strong password, there's no email reset" — so
 * the user takes the password seriously. Phase C Mode 3 ships
 * without SMTP; the documented recovery is a `docker exec ...
 * reset-password <username>` (Batch 5).
 *
 * In `multi` mode the SignupView is also shown to additional users
 * who hit the page directly — the gate renders this when
 * `signupAllowed` is true regardless of whether any users exist.
 */

export function SignupView() {
  const { setAuthenticated, refresh, state } = useAuth();
  // Show "create your admin account" copy only when this is the
  // first user. In multi mode with existing users it just says
  // "create your account."
  const isFirstUser = state.kind === 'unauthenticated' ? !state.hasAnyUser : false;
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const onSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setError(null);
    setBusy(true);
    const result = await signup(username.trim(), password);
    setBusy(false);
    if (!result.ok) {
      setError(humanise(result.reason));
      return;
    }
    // The server set the session cookie; flip the gate optimistically
    // and re-probe so the rest of the auth state (mode, etc.) is fresh.
    const mode = state.kind === 'unauthenticated' ? state.mode : 'single';
    setAuthenticated(result.user, mode);
    void refresh();
  };

  return (
    <form className="auth-card" onSubmit={onSubmit} data-testid="auth-signup">
      <h1 className="auth-card__title">
        {isFirstUser ? 'Welcome. Create your account.' : 'Create your account.'}
      </h1>
      {isFirstUser && (
        <p className="auth-card__hint">
          This account is the admin. Choose a strong password — there is no email reset.
        </p>
      )}

      <label className="auth-card__field">
        <span>Username</span>
        <input
          type="text"
          autoComplete="username"
          autoFocus
          value={username}
          onChange={(e) => setUsername(e.target.value)}
          required
          minLength={2}
          maxLength={40}
          pattern="[A-Za-z0-9_.\\-]{2,40}"
          data-testid="auth-signup-username"
        />
      </label>

      <label className="auth-card__field">
        <span>Password</span>
        <input
          type="password"
          autoComplete="new-password"
          value={password}
          onChange={(e) => setPassword(e.target.value)}
          required
          minLength={8}
          data-testid="auth-signup-password"
        />
      </label>

      {error && (
        <div className="auth-card__error" role="alert" data-testid="auth-error">
          {error}
        </div>
      )}

      <button
        type="submit"
        className="auth-btn auth-btn--primary"
        disabled={busy}
        data-testid="auth-signup-submit"
      >
        {busy ? 'Creating account…' : 'Create account'}
      </button>
    </form>
  );
}

function humanise(reason: string): string {
  switch (reason) {
    case 'username-taken':
      return 'That username is already taken.';
    case 'invalid-username':
      return 'Username must be 2–40 chars: letters, digits, _ . -';
    case 'weak-password':
      return 'Password must be at least 8 characters.';
    case 'signup-closed':
      return 'Signup is closed on this server.';
    case 'personal-mode-disabled':
      return 'Personal accounts are disabled on this server.';
    default:
      return `Could not sign up (${reason}).`;
  }
}
