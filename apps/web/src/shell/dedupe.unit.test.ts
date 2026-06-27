import { strict as assert } from 'node:assert';
import { test } from 'node:test';
import { dedupeRows, type CellRow } from './dedupe.ts';

const r = (...vals: Array<unknown>): CellRow => vals.map((v) => (v == null ? null : { v }));

test('removes duplicate rows across all columns, keeping the first occurrence', () => {
  const rows = [r('Ada', 90), r('Bob', 88), r('Ada', 90), r('Cleo', 75), r('Bob', 88)];
  const out = dedupeRows({ rows, hasHeaders: false, compareColumns: [0, 1] });
  assert.equal(out.removed, 2);
  assert.equal(out.remaining, 3);
  // Compacted to the top, originals preserved.
  assert.deepEqual(
    out.rows.map((row) => row.map((c) => c?.v ?? null)),
    [
      ['Ada', 90],
      ['Bob', 88],
      ['Cleo', 75],
      [null, null],
      [null, null],
    ],
  );
});

test('keeps the header row and never compares it', () => {
  const rows = [r('Name', 'Score'), r('Ada', 90), r('Ada', 90)];
  const out = dedupeRows({ rows, hasHeaders: true, compareColumns: [0, 1] });
  assert.equal(out.removed, 1);
  assert.equal(out.remaining, 1);
  assert.deepEqual(
    out.rows[0].map((c) => c?.v),
    ['Name', 'Score'],
  );
  assert.deepEqual(
    out.rows[1].map((c) => c?.v),
    ['Ada', 90],
  );
  assert.deepEqual(
    out.rows[2].map((c) => c?.v ?? null),
    [null, null],
  );
});

test('compares only the chosen columns', () => {
  // Same name, different score → duplicates only when comparing column 0.
  const rows = [r('Ada', 90), r('Ada', 12), r('Bob', 5)];
  const onCol0 = dedupeRows({ rows, hasHeaders: false, compareColumns: [0] });
  assert.equal(onCol0.removed, 1); // second Ada dropped
  const onBoth = dedupeRows({ rows, hasHeaders: false, compareColumns: [0, 1] });
  assert.equal(onBoth.removed, 0); // (Ada,90) ≠ (Ada,12)
});

test('preserves formulas and styles on kept rows', () => {
  const rows: CellRow[] = [
    [
      { v: 1, s: 'styleA' },
      { f: '=A1*2', v: 2 },
    ],
    [
      { v: 1, s: 'styleA' },
      { f: '=A1*2', v: 2 },
    ], // duplicate
  ];
  const out = dedupeRows({ rows, hasHeaders: false, compareColumns: [0, 1] });
  assert.equal(out.removed, 1);
  assert.equal(out.rows[0][0]?.s, 'styleA');
  assert.equal(out.rows[0][1]?.f, '=A1*2');
});

test('no comparison columns removes nothing', () => {
  const rows = [r('Ada', 90), r('Ada', 90)];
  const out = dedupeRows({ rows, hasHeaders: false, compareColumns: [] });
  assert.equal(out.removed, 0);
  assert.equal(out.remaining, 2);
});

test('all-unique data is unchanged', () => {
  const rows = [r('a'), r('b'), r('c')];
  const out = dedupeRows({ rows, hasHeaders: false, compareColumns: [0] });
  assert.equal(out.removed, 0);
  assert.deepEqual(
    out.rows.map((row) => row[0]?.v),
    ['a', 'b', 'c'],
  );
});
