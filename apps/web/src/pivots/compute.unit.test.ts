import { strict as assert } from 'node:assert';
import { test } from 'node:test';
import { computePivot, dateGroupKey, type SourceMatrix } from './compute.ts';
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const model = (m: Record<string, unknown>): any => m;

test('dateGroupKey buckets an Excel serial by year', () => {
  // 45658 = 2025-01-01 (the importer stores dates as serials).
  assert.equal(dateGroupKey(45658, 'year'), '2025');
  assert.equal(dateGroupKey(45658, 'quarter'), '2025-Q1');
  assert.equal(dateGroupKey(45658, 'month'), '2025-01');
});

test('dateGroupKey buckets a date string (DATE() formula shape)', () => {
  // A =DATE() cell surfaces as "2025/06/20" in this app.
  assert.equal(dateGroupKey('2025/06/20', 'year'), '2025');
  assert.equal(dateGroupKey('2025/06/20', 'quarter'), '2025-Q2');
  assert.equal(dateGroupKey('2025/06/20', 'month'), '2025-06');
  assert.equal(dateGroupKey('2025-03-15', 'month'), '2025-03');
  assert.equal(dateGroupKey('2025/12/31', 'quarter'), '2025-Q4');
});

test('quarter boundaries', () => {
  assert.equal(dateGroupKey('2025-01-01', 'quarter'), '2025-Q1');
  assert.equal(dateGroupKey('2025-04-01', 'quarter'), '2025-Q2');
  assert.equal(dateGroupKey('2025-07-01', 'quarter'), '2025-Q3');
  assert.equal(dateGroupKey('2025-10-01', 'quarter'), '2025-Q4');
});

test('non-date values fall back to the raw key', () => {
  assert.equal(dateGroupKey('East', 'year'), 'East');
  assert.equal(dateGroupKey('notadate', 'quarter'), 'notadate');
  assert.equal(dateGroupKey(null, 'year'), '');
});

test("'none' returns the raw key unchanged", () => {
  assert.equal(dateGroupKey(45658, 'none'), '45658');
  assert.equal(dateGroupKey('East', 'none'), 'East');
});

test('grouped keys sort chronologically as strings', () => {
  const keys = ['2025/03/01', '2025/01/01', '2026/01/01'].map((d) => dateGroupKey(d, 'month'));
  assert.deepEqual([...keys].sort(), ['2025-01', '2025-03', '2026-01']);
});

// ── Nested column fields (cross-tab with 2 column fields) ────────────────────
const NESTED_SRC: SourceMatrix = {
  headers: ['Region', 'Quarter', 'Month', 'Sales'],
  records: [
    ['East', 'Q1', 'Jan', 10],
    ['East', 'Q1', 'Feb', 20],
    ['East', 'Q2', 'Apr', 30],
    ['West', 'Q1', 'Jan', 40],
  ],
};

test('nested column fields produce N-level headers + tuple-sliced values', () => {
  const { grid, colMeta } = computePivot(
    NESTED_SRC,
    model({
      rows: [{ column: 0 }], // Region
      cols: [{ column: 1 }, { column: 2 }], // Quarter, then Month
      values: [{ column: 3, agg: 'sum' }],
    }),
  );
  // Tuples sort level-by-level: (Q1,Feb) < (Q1,Jan) < (Q2,Apr).
  assert.deepEqual(grid[0], ['Region', 'Q1', '', 'Q2', 'Grand Total']); // Quarter span
  assert.deepEqual(grid[1], ['', 'Feb', 'Jan', 'Apr', '']); // Month level
  assert.deepEqual(grid[2], ['East', 20, 10, 30, 60]);
  assert.deepEqual(grid[3], ['West', 0, 40, 0, 40]); // no East-only tuples
  assert.deepEqual(grid[4], ['Grand Total', 20, 50, 30, 100]);
  // colMeta carries the full key tuple for drill-down.
  assert.deepEqual(colMeta?.[1], { kind: 'value', colKeys: ['Q1', 'Feb'], valueIndex: 0 });
  assert.deepEqual(colMeta?.[3], { kind: 'value', colKeys: ['Q2', 'Apr'], valueIndex: 0 });
  assert.equal(colMeta?.[4]?.kind, 'grand-total');
});

test('a single column field is unchanged by the nesting generalisation', () => {
  const src: SourceMatrix = {
    headers: ['Region', 'Quarter', 'Sales'],
    records: [
      ['East', 'Q1', 100],
      ['East', 'Q2', 200],
      ['West', 'Q1', 150],
    ],
  };
  const { grid, colMeta } = computePivot(
    src,
    model({ rows: [{ column: 0 }], cols: [{ column: 1 }], values: [{ column: 2, agg: 'sum' }] }),
  );
  // One header row, single colKey per column (identical to the pre-nesting path).
  assert.deepEqual(grid[0], ['Region', 'Q1', 'Q2', 'Grand Total']);
  assert.deepEqual(grid[1], ['East', 100, 200, 300]);
  assert.deepEqual(grid[2], ['West', 150, 0, 150]);
  assert.deepEqual(grid[3], ['Grand Total', 250, 200, 450]);
  assert.deepEqual(colMeta?.[1], { kind: 'value', colKeys: ['Q1'], valueIndex: 0 });
});
