import { useState, type FormEvent } from 'react';
import { adminApi, AdminApiError, setToken } from './api';

interface Props {
  onLoggedIn: () => void;
}

export function AdminLogin({ onLoggedIn }: Props) {
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const submit = async (e: FormEvent) => {
    e.preventDefault();
    setBusy(true);
    setError(null);
    try {
      const { token } = await adminApi.login(username, password);
      setToken(token);
      onLoggedIn();
    } catch (err) {
      if (err instanceof AdminApiError && err.status === 401) {
        setError('Incorrect username or password.');
      } else {
        setError(err instanceof Error ? err.message : 'Login failed.');
      }
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className="admin-bootstrap">
      <form className="admin-card admin-login" onSubmit={submit}>
        <h1>Casual Sheets · Admin</h1>
        <p className="admin-card__hint">
          Sign in with the operator credentials set on the server via{' '}
          <code>CASUAL_ADMIN_USERNAME</code> + <code>CASUAL_ADMIN_PASSWORD</code>.
        </p>
        <label className="admin-field">
          <span>Username</span>
          <input
            type="text"
            value={username}
            onChange={(e) => setUsername(e.target.value)}
            autoFocus
            autoComplete="username"
            required
          />
        </label>
        <label className="admin-field">
          <span>Password</span>
          <input
            type="password"
            value={password}
            onChange={(e) => setPassword(e.target.value)}
            autoComplete="current-password"
            required
          />
        </label>
        {error && (
          <div className="admin-error" role="alert">
            {error}
          </div>
        )}
        <button type="submit" className="admin-btn admin-btn--primary" disabled={busy}>
          {busy ? 'Signing in…' : 'Sign in'}
        </button>
      </form>
    </div>
  );
}
