import type { PersonalUser } from './types';

/**
 * Profile + avatar HTTP layer. Distinct from `api.ts` (which handles
 * the auth-state lifecycle: signup / login / logout / status) so
 * components that only need the editable profile don't import the
 * session-mutation helpers.
 */

export type UserProfile = {
  displayName: string | null;
  email: string | null;
  timezone: string;
  hasAvatar: boolean;
  preferences: Record<string, unknown>;
};

const COMMON: RequestInit = { credentials: 'include' };

export async function fetchProfile(): Promise<{ user: PersonalUser; profile: UserProfile }> {
  const res = await fetch('/auth/profile', { ...COMMON });
  if (!res.ok) throw new Error(`fetch profile: HTTP ${res.status}`);
  return (await res.json()) as { user: PersonalUser; profile: UserProfile };
}

export async function patchProfile(
  patch: Partial<Pick<UserProfile, 'displayName' | 'email' | 'timezone' | 'preferences'>>,
): Promise<{ ok: boolean; profile?: UserProfile; reason?: string }> {
  const res = await fetch('/auth/profile', {
    ...COMMON,
    method: 'PATCH',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify(patch),
  });
  if (res.ok) {
    const body = (await res.json()) as { profile: UserProfile };
    return { ok: true, profile: body.profile };
  }
  return { ok: false, reason: await readErrorReason(res) };
}

export async function uploadAvatar(file: Blob): Promise<{ ok: boolean; reason?: string }> {
  const form = new FormData();
  form.append('avatar', file);
  const res = await fetch('/auth/profile/avatar', {
    ...COMMON,
    method: 'POST',
    body: form,
  });
  if (res.ok) return { ok: true };
  return { ok: false, reason: await readErrorReason(res) };
}

export async function deleteAvatar(): Promise<void> {
  await fetch('/auth/profile/avatar', { ...COMMON, method: 'DELETE' });
}

/** Stable URL the AccountMenu (and any future avatar UI) renders.
 *  Includes the `?id=` query so callers can fetch avatars of other
 *  users in multi mode. The cache-buster `?t=` lets a fresh upload
 *  invalidate the previous image without waiting for the 2-minute
 *  HTTP cache. */
export function avatarUrl(userId: number, cacheBuster?: number): string {
  const t = cacheBuster ?? '';
  return `/auth/profile/avatar?id=${userId}${t ? `&t=${t}` : ''}`;
}

async function readErrorReason(res: Response): Promise<string> {
  try {
    const body = (await res.json()) as { error?: string };
    if (body.error) return body.error;
  } catch {
    /* not json */
  }
  return `HTTP ${res.status}`;
}
