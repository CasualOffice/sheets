/**
 * Admin REST client.
 *
 * Talks to `/api/admin/*`. Carries the admin-role JWT in
 * `Authorization: Bearer` for every authed call. The token is stashed
 * in localStorage under `casual.admin.token` after a successful login
 * + restored on next visit; logout clears it.
 *
 * Failed requests throw an `AdminApiError` carrying the HTTP status
 * + the server's JSON `error` code; the panel surfaces them as toasts.
 */

import type { AdminConfig } from './types';

const TOKEN_KEY = 'casual.admin.token';

export class AdminApiError extends Error {
  constructor(
    message: string,
    public readonly status: number,
    public readonly code?: string,
  ) {
    super(message);
    this.name = 'AdminApiError';
  }
}

export function getToken(): string | null {
  try {
    return localStorage.getItem(TOKEN_KEY);
  } catch {
    return null;
  }
}

export function setToken(tok: string | null): void {
  try {
    if (tok) localStorage.setItem(TOKEN_KEY, tok);
    else localStorage.removeItem(TOKEN_KEY);
  } catch {
    /* localStorage might be disabled — degrade silently */
  }
}

async function request<T>(
  path: string,
  init: RequestInit & { authed?: boolean } = {},
): Promise<T> {
  const headers: Record<string, string> = {
    'content-type': 'application/json',
    ...(init.headers as Record<string, string> | undefined),
  };
  if (init.authed !== false) {
    const tok = getToken();
    if (tok) headers.authorization = `Bearer ${tok}`;
  }
  const res = await fetch(path, { ...init, headers });
  let body: unknown;
  const ct = res.headers.get('content-type') ?? '';
  if (ct.includes('application/json')) body = await res.json().catch(() => undefined);
  else body = await res.text().catch(() => undefined);
  if (!res.ok) {
    const code = (body as { error?: string } | undefined)?.error;
    throw new AdminApiError(
      code ?? `${res.status} ${res.statusText}`,
      res.status,
      code,
    );
  }
  return body as T;
}

export const adminApi = {
  status: () => request<{ configured: boolean }>('/api/admin/status', { authed: false }),

  login: (username: string, password: string) =>
    request<{ token: string; ttl_seconds: number; username: string }>(
      '/api/admin/login',
      {
        method: 'POST',
        body: JSON.stringify({ username, password }),
        authed: false,
      },
    ),

  getConfig: () => request<AdminConfig>('/api/admin/config'),

  putConfig: (patch: Partial<AdminConfig>) =>
    request<AdminConfig>('/api/admin/config', {
      method: 'PUT',
      body: JSON.stringify(patch),
    }),

  me: () => request<{ anonymous: boolean; role: string; sub?: string }>('/api/me'),
};
