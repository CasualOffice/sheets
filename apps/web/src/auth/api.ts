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

import type { AuthState, PersonalUser } from './types';

/**
 * Thin HTTP layer for `/auth/*`. Used by the gate + account modal.
 * Every method credentials-includes so the session cookie tags
 * along. Tolerant of network / mode-off failures — the gate uses
 * the discriminated `AuthState` to render the right surface.
 */

const COMMON: RequestInit = { credentials: 'include' };

export type StatusBody = {
  mode: 'none' | 'single' | 'multi';
  signupAllowed: boolean;
  hasAnyUser: boolean;
  user: PersonalUser | null;
};

/**
 * Boot probe. Resolves to the `AuthState` the gate renders against.
 * Treats 404 (route not registered) and 503 (mode=none) and
 * network errors all as "personal auth disabled" so the GitHub
 * Pages build keeps working unchanged.
 */
export async function fetchStatus(signal?: AbortSignal): Promise<AuthState> {
  try {
    const res = await fetch('/auth/status', { ...COMMON, signal });
    if (res.status === 404 || res.status === 503) return { kind: 'disabled' };
    if (!res.ok) {
      return { kind: 'unreachable', message: `HTTP ${res.status}` };
    }
    const body = (await res.json()) as StatusBody;
    if (body.mode === 'none') return { kind: 'disabled' };
    if (body.user) {
      return { kind: 'authenticated', mode: body.mode, user: body.user };
    }
    return {
      kind: 'unauthenticated',
      mode: body.mode,
      signupAllowed: body.signupAllowed,
      hasAnyUser: body.hasAnyUser,
    };
  } catch (err) {
    if (err instanceof DOMException && err.name === 'AbortError') {
      // Caller cancelled — propagate so the effect can ignore the
      // result. Treating abort as "unreachable" would briefly flash
      // the offline view on every probe.
      throw err;
    }
    // No network = no server. Equivalent to mode=none for the gate.
    return { kind: 'disabled' };
  }
}

export type AuthResult = { ok: true; user: PersonalUser } | { ok: false; reason: string };

export async function signup(username: string, password: string): Promise<AuthResult> {
  const res = await fetch('/auth/signup', {
    ...COMMON,
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ username, password }),
  });
  if (res.ok) {
    const body = (await res.json()) as { user: PersonalUser };
    return { ok: true, user: body.user };
  }
  return { ok: false, reason: await readErrorReason(res) };
}

export async function login(username: string, password: string): Promise<AuthResult> {
  const res = await fetch('/auth/login', {
    ...COMMON,
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ username, password }),
  });
  if (res.ok) {
    const body = (await res.json()) as { user: PersonalUser };
    return { ok: true, user: body.user };
  }
  return { ok: false, reason: await readErrorReason(res) };
}

export async function logout(): Promise<void> {
  await fetch('/auth/logout', { ...COMMON, method: 'POST' });
}

export async function changePassword(
  currentPassword: string,
  newPassword: string,
): Promise<{ ok: boolean; reason?: string }> {
  const res = await fetch('/auth/change-password', {
    ...COMMON,
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ currentPassword, newPassword }),
  });
  if (res.ok) return { ok: true };
  return { ok: false, reason: await readErrorReason(res) };
}

export async function deleteAccount(): Promise<{ ok: boolean; reason?: string }> {
  const res = await fetch('/auth/delete-account', { ...COMMON, method: 'POST' });
  if (res.ok) return { ok: true };
  return { ok: false, reason: await readErrorReason(res) };
}

async function readErrorReason(res: Response): Promise<string> {
  try {
    const body = (await res.json()) as { error?: string };
    if (body.error) return body.error;
  } catch {
    /* not json — fall through */
  }
  return `HTTP ${res.status}`;
}
