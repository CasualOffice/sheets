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
 * Unit coverage for the secure share-link join helpers (sharing-model
 * §6.1, client side). Pure parsing + key composition — no React / WS,
 * so it runs under `node:test` without the editor graph.
 *
 * Run with `pnpm test:unit`.
 */
import { strict as assert } from 'node:assert';
import { test } from 'node:test';

import { parseShareMeta, shareMetaUrl, sharePasswordKey } from './share-meta.js';

test('parseShareMeta: a valid password-protected meta narrows to the typed shape', () => {
  const meta = parseShareMeta({ valid: true, role: 'edit', hasPassword: true, roomId: 'room-1' });
  assert.deepEqual(meta, { valid: true, role: 'edit', hasPassword: true, roomId: 'room-1' });
});

test('parseShareMeta: a valid no-password meta narrows with hasPassword=false', () => {
  const meta = parseShareMeta({ valid: true, role: 'view', hasPassword: false, roomId: 'r2' });
  assert.deepEqual(meta, { valid: true, role: 'view', hasPassword: false, roomId: 'r2' });
});

test('parseShareMeta: { valid: false } passes through (dead/unknown/expired)', () => {
  assert.deepEqual(parseShareMeta({ valid: false }), { valid: false });
});

test('parseShareMeta: a truncated valid body (missing fields) is rejected as null', () => {
  // Defends against a partial response being read as "no password
  // needed" — a valid:true body MUST carry all three fields.
  assert.equal(parseShareMeta({ valid: true }), null);
  assert.equal(parseShareMeta({ valid: true, role: 'edit' }), null);
  assert.equal(parseShareMeta({ valid: true, role: 'edit', roomId: 'r' }), null);
  // Wrong types are rejected too.
  assert.equal(
    parseShareMeta({ valid: true, role: 'edit', hasPassword: 'yes', roomId: 'r' }),
    null,
  );
});

test('parseShareMeta: non-object / malformed input → null', () => {
  for (const bad of [null, undefined, 42, 'nope', [], { valid: 'maybe' }, {}]) {
    assert.equal(parseShareMeta(bad), null, `expected null for ${JSON.stringify(bad)}`);
  }
});

test('shareMetaUrl: builds the public /meta path and encodes the token', () => {
  assert.equal(shareMetaUrl('tok123'), '/files/shares/link/tok123/meta');
  assert.equal(shareMetaUrl('a/b+c'), '/files/shares/link/a%2Fb%2Bc/meta');
});

test('sharePasswordKey: namespaced distinctly from the room-password key', () => {
  assert.equal(sharePasswordKey('tok123'), 'casual.collab.sp.tok123');
  // Must NOT collide with the room-password `casual.collab.pw.<id>` key.
  assert.notEqual(sharePasswordKey('x'), 'casual.collab.pw.x');
});
