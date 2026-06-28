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

import { colorSort, distinctColors, type SortRow } from './color-sort.js';

const RED = '#ff0000';
const WHITE = '#ffffff';

/** rows keyed by a single column; payload in col 1 tracks identity. */
function rows(keys: string[]): SortRow[] {
  return keys.map((k, i) => [{ v: k }, { v: `p${i}` }]);
}
const payloadsOf = (rs: SortRow[]) => rs.map((r) => (r[1] as { v: string }).v);

test('brings matching-colour rows to the top, stable for the rest', () => {
  const out = colorSort({
    rows: rows(['a', 'b', 'c', 'd']),
    colors: [WHITE, RED, WHITE, RED],
    hasHeaders: false,
    targetColor: RED,
    onTop: true,
  });
  // Red rows (b=p1, d=p3) first in order, then the rest (a=p0, c=p2).
  assert.deepEqual(payloadsOf(out.rows), ['p1', 'p3', 'p0', 'p2']);
});

test('on bottom sinks matching rows, rest stays on top', () => {
  const out = colorSort({
    rows: rows(['a', 'b', 'c', 'd']),
    colors: [WHITE, RED, WHITE, RED],
    hasHeaders: false,
    targetColor: RED,
    onTop: false,
  });
  assert.deepEqual(payloadsOf(out.rows), ['p0', 'p2', 'p1', 'p3']);
});

test('keeps the header row in place and aligns colours past it', () => {
  const data: SortRow[] = [[{ v: 'H' }, { v: 'h' }], ...rows(['a', 'b', 'c'])];
  const out = colorSort({
    rows: data,
    colors: [WHITE, WHITE, RED, WHITE], // header + 3 data rows
    hasHeaders: true,
    targetColor: RED,
    onTop: true,
  });
  assert.equal((out.rows[0][0] as { v: string }).v, 'H');
  // data was [a(p0,white), b(p1,red), c(p2,white)] → red first: b, a, c.
  assert.deepEqual(payloadsOf(out.rows.slice(1)), ['p1', 'p0', 'p2']);
});

test('is immutable + clones cells (output not aliased to input)', () => {
  const input = rows(['a', 'b']);
  const out = colorSort({
    rows: input,
    colors: [RED, WHITE],
    hasHeaders: false,
    targetColor: RED,
    onTop: true,
  });
  assert.notEqual(out.rows[0][0], input[0][0], 'cells are fresh objects');
});

test('no matching colour → order unchanged', () => {
  const out = colorSort({
    rows: rows(['a', 'b']),
    colors: [WHITE, WHITE],
    hasHeaders: false,
    targetColor: RED,
    onTop: true,
  });
  assert.deepEqual(payloadsOf(out.rows), ['p0', 'p1']);
});

test('distinctColors returns first-appearance order, skipping the header', () => {
  assert.deepEqual(distinctColors([WHITE, WHITE, RED, RED, WHITE], true), [WHITE, RED]);
  assert.deepEqual(distinctColors([RED, WHITE, RED], false), [RED, WHITE]);
});
