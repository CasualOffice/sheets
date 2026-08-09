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

import { strict as assert } from 'node:assert';
import { test } from 'node:test';
import type { IWorkbookData } from '@univerjs/core';
import {
  buildReplayPlan,
  COLLAB_LOG_PROTOCOL_VERSION,
  CollabProtocolError,
  mergeFrontier,
  recordId,
  type MutationRecord,
  type SnapshotRecord,
} from './replay-plan';

const workbook = { id: 'wb', name: 'Workbook', sheetOrder: [], sheets: {} } as IWorkbookData;

function op(c: string, s: number, id = 'sheet.mutation.set-range-values'): MutationRecord {
  return { v: COLLAB_LOG_PROTOCOL_VERSION, kind: 'op', c, s, t: s, id, p: { value: s } };
}

function snapshot(c: string, s: number, frontier: Record<string, number>): SnapshotRecord {
  return {
    v: COLLAB_LOG_PROTOCOL_VERSION,
    kind: 'snapshot',
    c,
    s,
    t: s,
    wb: workbook,
    frontier,
  };
}

test('stable ids preserve a concurrent insert before an already-seen position', () => {
  const a = op('a', 0);
  const b = op('b', 0);
  const first = buildReplayPlan([b]);
  const merged = buildReplayPlan([a, b]);
  assert.deepEqual(first.ids, ['b:0']);
  assert.deepEqual(merged.ids, ['a:0', 'b:0']);
  assert.equal(recordId(a), 'a:0');
});

test('snapshot frontier filters covered operations regardless of physical array position', () => {
  const base = snapshot('a', 2, { a: 1, b: 4 });
  const plan = buildReplayPlan([op('b', 3), op('c', 0), base, op('a', 1), op('b', 5)]);
  assert.equal(plan.snapshot, base);
  assert.deepEqual(plan.mutations.map(recordId), ['c:0', 'b:5']);
  assert.deepEqual(plan.ids, ['a:2', 'c:0', 'b:5']);
});

test('later snapshot dominates an older snapshot through its frontier', () => {
  const older = snapshot('a', 2, { a: 1 });
  const newer = snapshot('b', 4, { a: 2, b: 3 });
  const plan = buildReplayPlan([older, newer, op('c', 0)]);
  assert.equal(plan.snapshot, newer);
  assert.deepEqual(plan.ids, ['b:4', 'c:0']);
});

test('concurrent snapshots fail closed instead of selecting a divergent base', () => {
  assert.throws(
    () => buildReplayPlan([snapshot('a', 2, { a: 1 }), snapshot('b', 2, { b: 1 })]),
    (err: unknown) =>
      err instanceof CollabProtocolError && err.message.includes('concurrent compaction'),
  );
});

test('snapshot dominance includes the full inherited frontier', () => {
  const older = snapshot('a', 3, { a: 2, offline: 8 });
  // Merely covering the older snapshot record is insufficient: a malformed
  // descendant that drops `offline:8` could otherwise erase those operations.
  const incomplete = snapshot('b', 2, { a: 3, b: 1 });
  assert.throws(
    () => buildReplayPlan([older, incomplete]),
    (err: unknown) =>
      err instanceof CollabProtocolError && err.message.includes('concurrent compaction'),
  );
});

test('legacy, duplicate, and malformed records fail closed', () => {
  assert.throws(() => buildReplayPlan([{ c: 'legacy', t: 1, id: 'x', p: {} }]));
  assert.throws(() => buildReplayPlan([op('a', 0), op('a', 0)]));
  assert.throws(() => buildReplayPlan([{ ...op('a', 0), kind: 'future-kind' }]));
});

test('frontier merge retains prior compaction ancestry and advances current records', () => {
  assert.deepEqual(mergeFrontier({ a: 7 }, [op('a', 4), op('b', 2), op('b', 5)]), {
    a: 7,
    b: 5,
  });
});
