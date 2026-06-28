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

import {
  createContext,
  useCallback,
  useContext,
  useEffect,
  useMemo,
  useState,
  type ReactNode,
} from 'react';
import { fetchStatus } from './api';
import type { AuthState, PersonalUser } from './types';

/**
 * Reactive `AuthState` shared by the gate, the title-bar account
 * menu, the account modal, and the `PersonalFileSource` (Batch 4).
 *
 * One boot probe runs on mount. After signup / login / logout, call
 * `refresh()` to re-fetch the status — that flips the state through
 * the discriminated kinds and re-renders consumers.
 *
 * The gate is the only place that conditionally renders based on
 * `state.kind`. Other consumers just want the live `PersonalUser`
 * (or `null` when not signed in / mode is off).
 */

type AuthCtx = {
  state: AuthState;
  refresh: () => Promise<void>;
  /** Optimistic local set — called from the signup / login views
   *  after a successful POST so the UI doesn't flicker through
   *  `loading` while the next `/auth/status` lands. */
  setAuthenticated: (user: PersonalUser, mode: 'single' | 'multi') => void;
  /** Optimistic local clear — called from logout / delete-account so
   *  the gate paints the unauthenticated surface before the next
   *  probe. */
  setUnauthenticated: () => void;
};

const AuthContext = createContext<AuthCtx | null>(null);

export function AuthProvider({ children }: { children: ReactNode }) {
  const [state, setState] = useState<AuthState>({ kind: 'loading' });

  const refresh = useCallback(async () => {
    try {
      const next = await fetchStatus();
      setState(next);
    } catch {
      // Aborted — leave the prior state in place.
    }
  }, []);

  useEffect(() => {
    const ctrl = new AbortController();
    void fetchStatus(ctrl.signal).then(
      (next) => setState(next),
      () => {
        /* aborted */
      },
    );
    return () => ctrl.abort();
  }, []);

  const setAuthenticated = useCallback((user: PersonalUser, mode: 'single' | 'multi') => {
    setState({ kind: 'authenticated', user, mode });
  }, []);

  const setUnauthenticated = useCallback(() => {
    setState((prev) =>
      prev.kind === 'authenticated'
        ? {
            kind: 'unauthenticated',
            mode: prev.mode,
            signupAllowed: prev.mode === 'multi',
            hasAnyUser: true,
          }
        : prev,
    );
  }, []);

  const value = useMemo<AuthCtx>(
    () => ({ state, refresh, setAuthenticated, setUnauthenticated }),
    [state, refresh, setAuthenticated, setUnauthenticated],
  );
  return <AuthContext.Provider value={value}>{children}</AuthContext.Provider>;
}

export function useAuth(): AuthCtx {
  const ctx = useContext(AuthContext);
  if (!ctx) throw new Error('useAuth must be used inside <AuthProvider>');
  return ctx;
}

/** Convenience — the signed-in user, or null. */
export function useCurrentUser(): PersonalUser | null {
  const { state } = useAuth();
  return state.kind === 'authenticated' ? state.user : null;
}
