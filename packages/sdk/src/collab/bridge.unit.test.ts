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
 * Pure-function unit tests for the bridge helpers. Runs under
 * `node --import tsx` via `node:test`; no extra test runner installed.
 *
 * Stateful bridge logic (Yjs observers, command-service hooks, compaction)
 * is covered by tests/e2e/coedit-*.spec.ts. This file only covers the
 * pure stuff that doesn't need a Univer instance.
 *
 * Run with: `pnpm test:unit`
 */
import { strict as assert } from 'node:assert';
import { test } from 'node:test';

import { deepRewriteUnitId } from './bridge-helpers';

test('deepRewriteUnitId rewrites top-level unitId', () => {
  const params = { unitId: 'sender-id', subUnitId: 'sheet-1', value: 'A' };
  const out = deepRewriteUnitId(params, 'local-id') as typeof params;
  assert.equal(out.unitId, 'local-id');
  assert.equal(out.subUnitId, 'sheet-1');
  assert.notEqual(out, params, 'should return a clone when changes happen');
});

test('deepRewriteUnitId rewrites nested unitId (range.unitId)', () => {
  const params = {
    range: { unitId: 'sender-id', startRow: 0, startColumn: 0 },
    rangeData: [{ unitId: 'sender-id', sr: 0 }],
  };
  const out = deepRewriteUnitId(params, 'local-id') as {
    range: { unitId: string };
    rangeData: Array<{ unitId: string }>;
  };
  assert.equal(out.range.unitId, 'local-id');
  assert.equal(out.rangeData[0].unitId, 'local-id');
});

test('deepRewriteUnitId rewrites unitId inside array entries', () => {
  const params = {
    moves: [
      { source: { unitId: 'sender-id' }, target: { unitId: 'sender-id' } },
      { source: { unitId: 'sender-id' } },
    ],
  };
  const out = deepRewriteUnitId(params, 'local-id') as {
    moves: Array<{ source: { unitId: string }; target?: { unitId: string } }>;
  };
  assert.equal(out.moves[0].source.unitId, 'local-id');
  assert.equal(out.moves[0].target?.unitId, 'local-id');
  assert.equal(out.moves[1].source.unitId, 'local-id');
});

test('deepRewriteUnitId returns input unchanged when no unitId present', () => {
  const params = { subUnitId: 'sheet-1', cellValue: { 0: { 0: { v: 1 } } } };
  const out = deepRewriteUnitId(params, 'local-id');
  assert.strictEqual(out, params, 'should return same reference when nothing changes');
});

test('deepRewriteUnitId leaves unitId unchanged when already local', () => {
  const params = { unitId: 'local-id', other: 1 };
  const out = deepRewriteUnitId(params, 'local-id');
  assert.strictEqual(out, params);
});

test('deepRewriteUnitId does not descend into class instances', () => {
  class Custom {
    unitId = 'sender-id';
  }
  const inst = new Custom();
  const params = { wrap: inst };
  const out = deepRewriteUnitId(params, 'local-id') as { wrap: Custom };
  // Class instances are not plain objects, so the walker stops at the
  // boundary — the unitId stays as-is. Univer mutation params are
  // required to be JSON-friendly plain objects; anything carrying a
  // class instance is out of scope for the bridge.
  assert.equal(out.wrap.unitId, 'sender-id');
});

test('deepRewriteUnitId handles deeply nested unitId chains', () => {
  const params = {
    a: { b: { c: { d: { unitId: 'sender-id' } } } },
  };
  const out = deepRewriteUnitId(params, 'local-id') as {
    a: { b: { c: { d: { unitId: string } } } };
  };
  assert.equal(out.a.b.c.d.unitId, 'local-id');
});

test('deepRewriteUnitId preserves null and primitive leaves', () => {
  const params = {
    unitId: 'sender-id',
    nullField: null,
    boolField: true,
    numField: 42,
    strField: 'hello',
  };
  const out = deepRewriteUnitId(params, 'local-id') as Record<string, unknown>;
  assert.equal(out.unitId, 'local-id');
  assert.equal(out.nullField, null);
  assert.equal(out.boolField, true);
  assert.equal(out.numField, 42);
  assert.equal(out.strField, 'hello');
});

/* ── rewriteJson1OpPathUnitId — Stream F1 drawing-sync fix ───────── */

import { rewriteJson1OpPathUnitId } from './bridge-helpers';

test('rewriteJson1OpPathUnitId swaps leading unitId in a single JSONOp', () => {
  // Realistic shape from `sheet.mutation.set-drawing-apply`:
  // path-prefix...trailing-component. Trailing component is an
  // object describing the mutation (insert {i: ...}, remove {r: 0},
  // edit {ed: ...}).
  const op = ['owner-wb', 'sheet-1', 'data', 'drawing-7', { i: { drawingId: 'drawing-7' } }];
  const out = rewriteJson1OpPathUnitId(op, 'owner-wb', 'joiner-wb') as unknown[];
  assert.equal(out[0], 'joiner-wb');
  assert.equal(out[1], 'sheet-1');
  assert.equal(out[2], 'data');
  assert.equal(out[3], 'drawing-7');
  // Trailing component should be preserved by reference (we shallow-copy).
  assert.deepEqual(out[4], { i: { drawingId: 'drawing-7' } });
});

test('rewriteJson1OpPathUnitId walks a JSONOpList (op of ops)', () => {
  const opList = [
    ['owner-wb', 'sheet-1', 'data', 'd-1', { i: { x: 1 } }],
    ['owner-wb', 'sheet-1', 'order', 0, { i: 'd-1' }],
  ];
  const out = rewriteJson1OpPathUnitId(opList, 'owner-wb', 'joiner-wb') as unknown[][];
  assert.equal(out.length, 2);
  assert.equal(out[0][0], 'joiner-wb');
  assert.equal(out[1][0], 'joiner-wb');
});

test('rewriteJson1OpPathUnitId returns input unchanged when no match at [0]', () => {
  // E.g. some other plugin's json1 op that doesn't lead with unitId.
  const op = ['something-else', 'whatever', { i: 1 }];
  const out = rewriteJson1OpPathUnitId(op, 'owner-wb', 'joiner-wb');
  assert.equal(out, op); // identity — no copy made when nothing to do
});

test('rewriteJson1OpPathUnitId is a no-op when old === new', () => {
  const op = ['same-wb', 'sheet-1', 'data', 'd-1', { i: { x: 1 } }];
  const out = rewriteJson1OpPathUnitId(op, 'same-wb', 'same-wb');
  assert.equal(out, op);
});

test('rewriteJson1OpPathUnitId is a no-op for non-array input', () => {
  assert.equal(rewriteJson1OpPathUnitId(null, 'a', 'b'), null);
  assert.equal(rewriteJson1OpPathUnitId(undefined, 'a', 'b'), undefined);
  assert.equal(rewriteJson1OpPathUnitId('string', 'a', 'b'), 'string');
  assert.equal(rewriteJson1OpPathUnitId(42, 'a', 'b'), 42);
});

test('rewriteJson1OpPathUnitId preserves a single-element op-list edge case', () => {
  // JSONOpList with one entry — detection key is "first element is array".
  const opList = [['owner-wb', 'sheet-1', 'data', 'd-1', { i: 1 }]];
  const out = rewriteJson1OpPathUnitId(opList, 'owner-wb', 'joiner-wb') as unknown[][];
  assert.equal(out.length, 1);
  assert.equal(out[0][0], 'joiner-wb');
});
