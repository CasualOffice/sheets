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
import { login } from './api';

/**
 * Login surface. Rendered when the gate's `unauthenticated` state
 * has `signupAllowed: false` (single mode + existing user, or multi
 * mode without an open signup flow). Wrong-password errors are
 * mapped to a deliberately vague "username or password is wrong" so
 * we don't leak which axis was incorrect.
 */

export function LoginView() {
  const { setAuthenticated, refresh, state } = useAuth();
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const onSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setError(null);
    setBusy(true);
    const result = await login(username.trim(), password);
    setBusy(false);
    if (!result.ok) {
      setError(humanise(result.reason));
      return;
    }
    const mode = state.kind === 'unauthenticated' ? state.mode : 'single';
    setAuthenticated(result.user, mode);
    void refresh();
  };

  return (
    <form className="auth-card" onSubmit={onSubmit} data-testid="auth-login">
      <h1 className="auth-card__title">Sign in</h1>
      <p className="auth-card__hint">Welcome back. Sign in to access your spreadsheets.</p>

      <label className="auth-card__field">
        <span>Username</span>
        <input
          type="text"
          autoComplete="username"
          autoFocus
          value={username}
          onChange={(e) => setUsername(e.target.value)}
          required
          data-testid="auth-login-username"
        />
      </label>

      <label className="auth-card__field">
        <span>Password</span>
        <input
          type="password"
          autoComplete="current-password"
          value={password}
          onChange={(e) => setPassword(e.target.value)}
          required
          data-testid="auth-login-password"
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
        data-testid="auth-login-submit"
      >
        {busy ? 'Signing in…' : 'Sign in'}
      </button>
    </form>
  );
}

function humanise(reason: string): string {
  switch (reason) {
    case 'invalid-credentials':
      return 'Username or password is wrong.';
    case 'personal-mode-disabled':
      return 'Personal accounts are disabled on this server.';
    default:
      return `Couldn’t sign in (${reason}).`;
  }
}
