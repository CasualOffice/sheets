/**
 * Unit tests for the collab user-set builder (T3.1 identity foundation). Runs
 * under `node --import tsx` via `node:test`; the pure builder is Univer-free so
 * it loads without a Univer instance.
 *
 * Run with: `pnpm test:unit`
 */
import { strict as assert } from 'node:assert';
import { test } from 'node:test';

import { buildCollabUserSet } from './user-manager-sync.pure';

const self = { userID: 'me', name: 'Me' };

test('current is self; peers become others', () => {
  const { current, others } = buildCollabUserSet(self, [
    { userId: 'a', name: 'Alice' },
    { userId: 'b', name: 'Bob' },
  ]);
  assert.deepEqual(current, self);
  assert.deepEqual(others, [
    { userID: 'a', name: 'Alice' },
    { userID: 'b', name: 'Bob' },
  ]);
});

test('drops peers without a userId', () => {
  const { others } = buildCollabUserSet(self, [
    { userId: '', name: 'NoId' },
    { userId: 'a', name: 'Alice' },
  ]);
  assert.deepEqual(others, [{ userID: 'a', name: 'Alice' }]);
});

test('de-dupes by userID; self wins and is never duplicated into others', () => {
  const { others } = buildCollabUserSet(self, [
    { userId: 'me', name: 'Me (other tab)' }, // same id as self → skipped
    { userId: 'a', name: 'Alice' },
    { userId: 'a', name: 'Alice 2' }, // duplicate id → first wins
  ]);
  assert.deepEqual(others, [{ userID: 'a', name: 'Alice' }]);
});
