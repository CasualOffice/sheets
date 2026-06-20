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
