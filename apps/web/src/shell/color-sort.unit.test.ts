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

import { colorMultiSort, distinctColors, type SortRow } from './color-sort.js';

const RED = '#ff0000';
const WHITE = '#ffffff';

/** rows keyed by a single column; payload in col 1 tracks identity. */
function rows(keys: string[]): SortRow[] {
  return keys.map((k, i) => [{ v: k }, { v: `p${i}` }]);
}
const payloadsOf = (rs: SortRow[]) => rs.map((r) => (r[1] as { v: string }).v);

test('distinctColors returns first-appearance order, skipping the header', () => {
  assert.deepEqual(distinctColors([WHITE, WHITE, RED, RED, WHITE], true), [WHITE, RED]);
  assert.deepEqual(distinctColors([RED, WHITE, RED], false), [RED, WHITE]);
});

const GREEN = '#00ff00';

test('colorMultiSort orders rows by the colour priority list', () => {
  // rows a..e coloured: green, red, white, red, green.
  const out = colorMultiSort({
    rows: rows(['a', 'b', 'c', 'd', 'e']),
    colors: [GREEN, RED, WHITE, RED, GREEN],
    hasHeaders: false,
    order: [RED, GREEN], // red first, then green; white (unlisted) last
  });
  // red (b=p1, d=p3), green (a=p0, e=p4), then white (c=p2).
  assert.deepEqual(payloadsOf(out.rows), ['p1', 'p3', 'p0', 'p4', 'p2']);
});

test('colorMultiSort keeps unlisted colours after, in original order', () => {
  const out = colorMultiSort({
    rows: rows(['a', 'b', 'c']),
    colors: [WHITE, RED, GREEN],
    hasHeaders: false,
    order: [RED], // only red prioritised; white + green stay in place after
  });
  assert.deepEqual(payloadsOf(out.rows), ['p1', 'p0', 'p2']);
});

test('colorMultiSort is stable for same-colour rows and clones cells', () => {
  const input = rows(['a', 'b', 'c']);
  const out = colorMultiSort({
    rows: input,
    colors: [RED, RED, RED],
    hasHeaders: false,
    order: [RED],
  });
  assert.deepEqual(payloadsOf(out.rows), ['p0', 'p1', 'p2']);
  assert.notEqual(out.rows[0][0], input[0][0], 'cells are fresh objects');
});

test('colorMultiSort keeps the header row in place', () => {
  const data: SortRow[] = [[{ v: 'H' }, { v: 'h' }], ...rows(['a', 'b'])];
  const out = colorMultiSort({
    rows: data,
    colors: [WHITE, WHITE, RED], // header + 2 data rows
    hasHeaders: true,
    order: [RED],
  });
  assert.equal((out.rows[0][0] as { v: string }).v, 'H');
  assert.deepEqual(payloadsOf(out.rows.slice(1)), ['p1', 'p0']);
});
