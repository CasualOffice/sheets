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

/**
 * Pure helpers for the secure share-link JOIN side (sharing-model §6.1).
 * Kept free of React / Univer / Hocuspocus imports so the parsing +
 * sessionStorage-key logic is unit-testable under `node:test` without the
 * editor graph. CollabDriver wires these into the connect flow.
 */

/**
 * Public share-link metadata, mirrored from the server's
 * `GET /files/shares/link/:token/meta` contract. A dead / unknown /
 * expired token returns `{ valid: false }` and NOTHING else — never the
 * role / room / hasPassword (so a probe can't learn a token existed).
 */
export type ShareMeta =
  | { valid: true; role: string; hasPassword: boolean; roomId: string }
  | { valid: false };

/**
 * Validate + narrow an untrusted /meta JSON body. Returns the typed
 * `ShareMeta` on a well-formed payload, or `null` when the shape is
 * unrecognised — callers treat null as "couldn't determine, let the WS
 * connect decide" rather than trusting a malformed response. A
 * `{ valid: true }` body MUST carry the role/room/hasPassword fields or
 * it's rejected (defends against a truncated/partial response being read
 * as "no password needed").
 */
export function parseShareMeta(body: unknown): ShareMeta | null {
  if (typeof body !== 'object' || body === null) return null;
  const b = body as Record<string, unknown>;
  if (b.valid === false) return { valid: false };
  if (b.valid === true) {
    if (
      typeof b.role !== 'string' ||
      typeof b.hasPassword !== 'boolean' ||
      typeof b.roomId !== 'string'
    ) {
      return null;
    }
    return { valid: true, role: b.role, hasPassword: b.hasPassword, roomId: b.roomId };
  }
  return null;
}

/** Public /meta URL for a token. */
export function shareMetaUrl(token: string): string {
  return `/files/shares/link/${encodeURIComponent(token)}/meta`;
}

/** sessionStorage key for a token's entered share password. Namespaced
 *  under `sp` to stay DISTINCT from the room-password `casual.collab.pw`
 *  key — they're separate gates with separate values. */
export function sharePasswordKey(token: string): string {
  return `casual.collab.sp.${token}`;
}
