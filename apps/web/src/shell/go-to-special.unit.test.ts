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
import { computeGoToSpecial, type CellMatrix } from './go-to-special.ts';

// A small sheet:
//      A(0)      B(1)      C(2)
// 1    "Name"    "Score"   (blank)
// 2    "Ada"     90        =B2*2
// 3    (blank)   (blank)   =SUM(B2:B2)
const SHEET: CellMatrix = {
  0: { 0: { v: 'Name' }, 1: { v: 'Score' } },
  1: { 0: { v: 'Ada' }, 1: { v: 90 }, 2: { f: '=B2*2', v: 180 } },
  2: { 2: { f: '=SUM(B2:B2)', v: 90 } },
};

const active = { row: 0, column: 0 };

test('constants selects literal cells, not formulas', () => {
  const { ranges } = computeGoToSpecial(SHEET, active, 'constants');
  const cells = ranges.flatMap((r) => {
    const out: string[] = [];
    for (let row = r.startRow; row <= r.endRow; row++)
      for (let col = r.startColumn; col <= r.endColumn; col++) out.push(`${row},${col}`);
    return out;
  });
  // A1,B1 (row 0) + A2,B2 (row 1) — but NOT C2/C3 (formulas).
  assert.deepEqual(cells.sort(), ['0,0', '0,1', '1,0', '1,1'].sort());
});

test('constants coalesces adjacent columns into one range per row', () => {
  const { ranges } = computeGoToSpecial(SHEET, active, 'constants');
  // Row 0: A1:B1 is a single run; row 1: A2:B2 a single run.
  const row0 = ranges.find((r) => r.startRow === 0);
  assert.deepEqual(row0, { startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 });
});

test('formulas selects only formula cells', () => {
  const { ranges } = computeGoToSpecial(SHEET, active, 'formulas');
  const cells = ranges.map((r) => `${r.startRow},${r.startColumn}`).sort();
  assert.deepEqual(cells, ['1,2', '2,2']); // C2, C3
});

test('blanks selects empty cells within the used range only', () => {
  const { ranges } = computeGoToSpecial(SHEET, active, 'blanks');
  const cells = ranges
    .flatMap((r) => {
      const out: string[] = [];
      for (let col = r.startColumn; col <= r.endColumn; col++) out.push(`${r.startRow},${col}`);
      return out;
    })
    .sort();
  // Used range is A1:C3. Blanks: C1 (0,2), A3 (2,0), B3 (2,1). Nothing past C3.
  assert.deepEqual(cells, ['0,2', '2,0', '2,1'].sort());
});

test('last cell is the bottom-right of the used range', () => {
  const { ranges } = computeGoToSpecial(SHEET, active, 'lastCell');
  assert.deepEqual(ranges, [{ startRow: 2, endRow: 2, startColumn: 2, endColumn: 2 }]);
});

test('current region is the contiguous block around the active cell', () => {
  const { ranges } = computeGoToSpecial(SHEET, active, 'currentRegion');
  // The whole A1:C3 block is 8-connected through the data.
  assert.deepEqual(ranges, [{ startRow: 0, endRow: 2, startColumn: 0, endColumn: 2 }]);
});

test('an empty sheet yields no cells for every criterion', () => {
  for (const crit of ['constants', 'formulas', 'blanks', 'lastCell', 'currentRegion'] as const) {
    assert.deepEqual(computeGoToSpecial({}, active, crit).ranges, [], crit);
  }
});

test('current region around an isolated island excludes a separate block', () => {
  // Two islands separated by a fully-empty row/col gap.
  const m: CellMatrix = {
    0: { 0: { v: 1 }, 1: { v: 2 } },
    1: { 0: { v: 3 }, 1: { v: 4 } },
    // gap at row 2 / col 2
    4: { 4: { v: 'far' } },
  };
  const { ranges } = computeGoToSpecial(m, { row: 0, column: 0 }, 'currentRegion');
  assert.deepEqual(ranges, [{ startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 }]);
});
