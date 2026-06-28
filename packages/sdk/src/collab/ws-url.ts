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
 * Pure WS-URL composition for attachCollab. Kept in its own module — free of
 * any `@univerjs` / `@hocuspocus` imports — so it's unit-testable under
 * `node:test` without dragging in the Univer ESM graph.
 */

/** Extra capability params carried on the WS upgrade for secure share
 *  links (sharing-model §6.1). */
export type WsUrlShare = {
  /** Secure share-link token. When present the server is authoritative
   *  for the role — it resolves the token to a role + a bound room, so
   *  the client must NOT also assert `role=` (the server ignores a
   *  client `?role=` when a token is present; sending one is just
   *  spoofable noise). */
  share?: string;
  /** Optional join password paired with a password-protected share
   *  token (`?sp=`). Distinct from the anonymous-room `?p=` password. */
  sp?: string;
};

/**
 * Build the room WS URL.
 *
 *   - No share token: `<server>?room=<id>[&p=<pw>]&role=<role>` — the
 *     anonymous-room path; byte-identical to before.
 *   - With a `share` token: `<server>?room=<id>[&p=<pw>]&share=<token>[&sp=<pw>]`.
 *     `role=` is deliberately omitted — the server resolves the role
 *     from the token, so shipping a spoofable client role is pointless.
 */
export function buildWsUrl(
  server: string,
  room: string,
  role: 'view' | 'write',
  password?: string,
  share?: WsUrlShare,
): string {
  const sep = server.includes('?') ? '&' : '?';
  const base =
    `${server}${sep}room=${encodeURIComponent(room)}` +
    `${password ? `&p=${encodeURIComponent(password)}` : ''}`;
  const token = share?.share;
  if (token) {
    return (
      base +
      `&share=${encodeURIComponent(token)}` +
      `${share?.sp ? `&sp=${encodeURIComponent(share.sp)}` : ''}`
    );
  }
  return base + `&role=${encodeURIComponent(role)}`;
}
