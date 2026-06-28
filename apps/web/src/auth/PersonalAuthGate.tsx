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

import { type ReactNode } from 'react';
import { useAuth } from './auth-context';
import { detectWopiContext } from '../file-source/wopi-file-source';
import { useRoute } from '../router';
import { LoginView } from './LoginView';
import { SignupView } from './SignupView';
import './auth.css';

/**
 * Wraps the app shell. On the GitHub Pages build (or any deploy
 * where `CASUAL_PERSONAL_MODE=none`) `/auth/status` 404s or returns
 * `mode: 'none'`; the gate renders nothing and the app boots
 * normally — Mode 1 / Mode 2 paths are unchanged.
 *
 * For `single|multi`:
 *
 *   - `loading`         — thin "Checking your account…" card so the
 *                         editor doesn't flash visible behind the gate.
 *   - `unreachable`     — small "Couldn't reach the server" card with
 *                         a Retry button. Doesn't boot the app into an
 *                         unknown state.
 *   - `unauthenticated` — `signupAllowed` → SignupView, else LoginView.
 *                         SignupView toggles to LoginView when an
 *                         account already exists ("multi" mode flow).
 *   - `authenticated`   — renders `children` (the rest of the app).
 */

export function PersonalAuthGate({ children }: { children: ReactNode }) {
  const { state, refresh } = useAuth();
  const route = useRoute();

  // Embedded-host (WOPI) deploys carry a URL access token and are
  // authenticated against the embedding host's identity, not the
  // self-host personal-mode users table. Skip the gate so the
  // editor mounts immediately and the WopiFileSource takes over.
  if (detectWopiContext()) return <>{children}</>;

  // Anonymous-collab room routes (`/r/<roomId>`, optionally with a
  // share token) skip the personal-auth gate entirely. The room is
  // joined as a guest via NamePill; gating it would force the joiner
  // to create or claim an account on the host's box, which isn't the
  // intent of share-for-edit. UX_AUDIT.md §2.8.
  if (route.kind === 'room') return <>{children}</>;

  if (state.kind === 'disabled') return <>{children}</>;
  if (state.kind === 'authenticated') return <>{children}</>;

  if (state.kind === 'loading') {
    return (
      <div className="auth-gate" data-testid="auth-gate-loading">
        <div className="auth-card auth-card--thin">
          <div className="auth-card__spinner" aria-hidden />
          <p className="auth-card__hint">Checking your account…</p>
        </div>
      </div>
    );
  }

  if (state.kind === 'unreachable') {
    return (
      <div className="auth-gate" data-testid="auth-gate-unreachable">
        <div className="auth-card">
          <h1 className="auth-card__title">Couldn’t reach the server</h1>
          <p className="auth-card__hint">{state.message}</p>
          <button
            type="button"
            className="auth-btn auth-btn--primary"
            onClick={() => void refresh()}
            data-testid="auth-retry"
          >
            Try again
          </button>
        </div>
      </div>
    );
  }

  // unauthenticated — pick signup vs login from `signupAllowed`.
  return (
    <div className="auth-gate" data-testid="auth-gate-unauthenticated">
      {state.signupAllowed ? <SignupView /> : <LoginView />}
    </div>
  );
}
