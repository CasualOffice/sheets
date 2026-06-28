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
 * Unit coverage for attachCollab's URL composition — the one bit of our own
 * logic that doesn't require a live WebSocket. The provider/bridge wiring is
 * exercised end-to-end by the reference host's `tests/e2e/coedit-*.spec.ts`
 * (which now drives the moved bridge through `@casualoffice/sheets/collab`).
 */
import { strict as assert } from 'node:assert';
import { test } from 'node:test';

import { buildWsUrl } from './ws-url';

test('buildWsUrl appends room + role with a leading ? on a bare server', () => {
  assert.equal(buildWsUrl('wss://h/yjs', 'room1', 'write'), 'wss://h/yjs?room=room1&role=write');
});

test('buildWsUrl uses & when the server URL already has a query', () => {
  assert.equal(
    buildWsUrl('wss://h/yjs?token=x', 'room1', 'view'),
    'wss://h/yjs?token=x&room=room1&role=view',
  );
});

test('buildWsUrl includes the password only when provided, before role', () => {
  assert.equal(
    buildWsUrl('wss://h/yjs', 'room1', 'write', 'p@ss/word'),
    'wss://h/yjs?room=room1&p=p%40ss%2Fword&role=write',
  );
});

test('buildWsUrl percent-encodes the room id', () => {
  assert.equal(
    buildWsUrl('wss://h/yjs', 'a b/c', 'write'),
    'wss://h/yjs?room=a%20b%2Fc&role=write',
  );
});

// ── Secure share-link tokens (sharing-model §6.1) ──────────────────────────

test('buildWsUrl with a share token forwards share= and OMITS role=', () => {
  // The server is authoritative for the role when a token is present, so the
  // spoofable client role must not be on the URL — even though `role` is
  // still passed positionally. A `write` here must NOT leak through.
  assert.equal(
    buildWsUrl('wss://h/yjs', 'room1', 'write', undefined, { share: 'tok123' }),
    'wss://h/yjs?room=room1&share=tok123',
  );
});

test('buildWsUrl with a share token forwards sp= (join password) when set', () => {
  assert.equal(
    buildWsUrl('wss://h/yjs', 'room1', 'view', undefined, { share: 'tok123', sp: 'p@ss/word' }),
    'wss://h/yjs?room=room1&share=tok123&sp=p%40ss%2Fword',
  );
});

test('buildWsUrl percent-encodes the share token', () => {
  assert.equal(
    buildWsUrl('wss://h/yjs', 'room1', 'view', undefined, { share: 'a b/c' }),
    'wss://h/yjs?room=room1&share=a%20b%2Fc',
  );
});

test('buildWsUrl keeps a room password alongside a share token', () => {
  assert.equal(
    buildWsUrl('wss://h/yjs', 'room1', 'write', 'roompw', { share: 'tok' }),
    'wss://h/yjs?room=room1&p=roompw&share=tok',
  );
});

test('buildWsUrl is byte-identical to the no-token form when share is undefined', () => {
  // The explicit-undefined call site (e.g. attachCollab forwarding an absent
  // opts.share) must produce exactly what the old 4-arg call produced.
  assert.equal(
    buildWsUrl('wss://h/yjs', 'room1', 'write', undefined, undefined),
    buildWsUrl('wss://h/yjs', 'room1', 'write'),
  );
  assert.equal(
    buildWsUrl('wss://h/yjs', 'room1', 'view', 'pw', undefined),
    buildWsUrl('wss://h/yjs', 'room1', 'view', 'pw'),
  );
});

test('buildWsUrl ignores an empty share token (falls back to role=)', () => {
  // A defensive case: `{ share: '' }` is not a real token, so we must not
  // emit `share=` with an empty value and drop the role.
  assert.equal(
    buildWsUrl('wss://h/yjs', 'room1', 'write', undefined, { share: '' }),
    'wss://h/yjs?room=room1&role=write',
  );
});
