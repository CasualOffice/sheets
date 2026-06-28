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

import {
  addWatches,
  cellA1,
  cellsInRect,
  removeWatch,
  watchKey,
  type Watch,
} from './watch-model.js';

const mk = (i: number) => `w${i}`;

test('cellA1 formats cell references', () => {
  assert.equal(cellA1(0, 0), 'A1');
  assert.equal(cellA1(4, 2), 'C5');
  assert.equal(cellA1(0, 26), 'AA1');
});

test('watchKey is stable per sheet+cell', () => {
  assert.equal(watchKey('s1', 2, 3), watchKey('s1', 2, 3));
  assert.notEqual(watchKey('s1', 2, 3), watchKey('s2', 2, 3));
});

test('addWatches appends and dedupes by sheet+cell', () => {
  let list: Watch[] = [];
  list = addWatches(list, [{ sheetId: 's1', sheetName: 'S', row: 0, col: 0 }], mk);
  assert.equal(list.length, 1);
  // Re-adding the same cell is a no-op; a new cell appends.
  list = addWatches(
    list,
    [
      { sheetId: 's1', sheetName: 'S', row: 0, col: 0 },
      { sheetId: 's1', sheetName: 'S', row: 1, col: 0 },
    ],
    mk,
  );
  assert.equal(list.length, 2);
  assert.deepEqual(
    list.map((w) => w.row),
    [0, 1],
  );
});

test('addWatches ids only the newly-added cells', () => {
  const list = addWatches(
    [],
    [
      { sheetId: 's1', sheetName: 'S', row: 0, col: 0 },
      { sheetId: 's1', sheetName: 'S', row: 0, col: 0 }, // dup, skipped
      { sheetId: 's1', sheetName: 'S', row: 1, col: 0 },
    ],
    mk,
  );
  assert.deepEqual(
    list.map((w) => w.id),
    ['w0', 'w1'],
  );
});

test('removeWatch drops by id', () => {
  const list: Watch[] = [
    { id: 'a', sheetId: 's', sheetName: 'S', row: 0, col: 0 },
    { id: 'b', sheetId: 's', sheetName: 'S', row: 1, col: 0 },
  ];
  assert.deepEqual(
    removeWatch(list, 'a').map((w) => w.id),
    ['b'],
  );
});

test('cellsInRect enumerates row-major and caps', () => {
  const cells = cellsInRect({ startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 });
  assert.deepEqual(cells, [
    { row: 0, col: 0 },
    { row: 0, col: 1 },
    { row: 1, col: 0 },
    { row: 1, col: 1 },
  ]);
  const big = cellsInRect({ startRow: 0, endRow: 999, startColumn: 0, endColumn: 999 }, 10);
  assert.equal(big.length, 10);
});
