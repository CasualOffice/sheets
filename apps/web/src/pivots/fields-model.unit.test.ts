import { strict as assert } from 'node:assert';
import { test } from 'node:test';

import {
  addFieldToZone,
  axisOf,
  hasValues,
  moveWithinZone,
  placedColumns,
  removeFieldFromZone,
  updateRowGrouping,
  updateValueField,
} from './fields-model.js';
import type { PivotModel } from './types.js';

function baseModel(): PivotModel {
  return {
    id: 'pt-1',
    sourceSheetId: 's1',
    source: { startRow: 0, endRow: 5, startColumn: 0, endColumn: 2 },
    targetSheetId: 's1',
    target: { row: 0, column: 4 },
    rows: [{ column: 0 }],
    cols: [],
    values: [{ column: 1, agg: 'sum', showAs: 'normal' }],
    filters: [],
  };
}

test('addFieldToZone places a column on Columns and strips it from Rows', () => {
  const m = addFieldToZone(baseModel(), 0, 'cols');
  assert.deepEqual(
    m.rows.map((r) => r.column),
    [],
    'column moved off Rows',
  );
  assert.deepEqual(
    m.cols.map((c) => c.column),
    [0],
  );
});

test('axes are mutually exclusive but Values is independent', () => {
  // Region (col 0) is on Rows; also aggregate it in Values (Count of Region).
  let m = addFieldToZone(baseModel(), 0, 'values', { defaultAgg: 'count' });
  assert.equal(m.rows.length, 1, 'still on Rows');
  assert.equal(m.values.length, 2, 'added to Values too');
  // Now move col 0 to Columns — leaves Rows, keeps the Values entry.
  m = addFieldToZone(m, 0, 'cols');
  assert.equal(m.rows.length, 0);
  assert.equal(m.cols.length, 1);
  assert.equal(
    m.values.filter((v) => v.column === 0).length,
    1,
    'Values entry survives the axis move',
  );
});

test('addFieldToZone is immutable — original model untouched', () => {
  const orig = baseModel();
  const before = JSON.stringify(orig);
  addFieldToZone(orig, 2, 'rows');
  assert.equal(JSON.stringify(orig), before, 'input not mutated');
});

test('Values may hold the same column twice; remove is by index', () => {
  let m = addFieldToZone(baseModel(), 1, 'values', { defaultAgg: 'average' });
  assert.equal(m.values.length, 2);
  m = removeFieldFromZone(m, 'values', 0);
  assert.equal(m.values.length, 1);
  assert.equal(m.values[0].agg, 'average', 'removed the right (first) entry');
});

test('filters get an allowedValues list when assigned', () => {
  const m = addFieldToZone(baseModel(), 2, 'filters', { allowedValues: ['a', 'b'] });
  assert.equal(m.filters?.length, 1);
  assert.deepEqual(m.filters?.[0], { column: 2, allowedValues: ['a', 'b'] });
});

test('moveWithinZone reorders and is a no-op for bad indices', () => {
  let m = baseModel();
  m = addFieldToZone(m, 2, 'rows'); // rows = [0, 2]
  m = moveWithinZone(m, 'rows', 1, 0);
  assert.deepEqual(
    m.rows.map((r) => r.column),
    [2, 0],
  );
  const same = moveWithinZone(m, 'rows', 5, 0);
  assert.equal(same, m, 'out-of-range move returns the same object');
});

test('updateValueField patches agg + showAs', () => {
  const m = updateValueField(baseModel(), 0, { agg: 'max', showAs: 'pctOfGrandTotal' });
  assert.equal(m.values[0].agg, 'max');
  assert.equal(m.values[0].showAs, 'pctOfGrandTotal');
});

test('updateRowGrouping sets and clears grouping', () => {
  let m = updateRowGrouping(baseModel(), 0, 'month');
  assert.equal(m.rows[0].grouping, 'month');
  m = updateRowGrouping(m, 0, 'none');
  assert.equal(m.rows[0].grouping, undefined, "'none' clears the field");
});

test('placedColumns + axisOf reflect placement', () => {
  let m = baseModel(); // rows[0], values[1]
  m = addFieldToZone(m, 2, 'filters', { allowedValues: ['x'] });
  assert.deepEqual(
    [...placedColumns(m)].sort((a, b) => a - b),
    [0, 1, 2],
  );
  assert.equal(axisOf(m, 0), 'rows');
  assert.equal(axisOf(m, 2), 'filters');
  assert.equal(axisOf(m, 1), null, 'a Values-only column is on no axis');
});

test('hasValues guards the last value field', () => {
  const empty = removeFieldFromZone(baseModel(), 'values', 0);
  assert.equal(hasValues(empty), false);
  assert.equal(hasValues(baseModel()), true);
});
