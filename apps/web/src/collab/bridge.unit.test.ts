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
