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
import { computeSubtotals, type SubtotalRow } from './subtotals.ts';

// Region, Amount — sorted by Region.
//   Region  Amount
//   East    10
//   East    20
//   West    30
const SHEET: SubtotalRow[] = [
  [{ v: 'Region' }, { v: 'Amount' }],
  [{ v: 'East' }, { v: 10 }],
  [{ v: 'East' }, { v: 20 }],
  [{ v: 'West' }, { v: 30 }],
];

test('inserts a subtotal row per group plus a grand total', () => {
  const out = computeSubtotals({
    rows: SHEET,
    groupCol: 0,
    functionCode: 9, // SUM
    subtotalCols: [1],
    startRow: 0,
    startCol: 0,
  });
  // 2 group subtotals + 1 grand = 3 inserted rows.
  assert.equal(out.groups, 2);
  assert.equal(out.insertedRows, 3);
  assert.equal(out.rows.length, 7);

  // Layout: header, East, East, East Total, West, West Total, Grand Total.
  const labels = out.rows.map((r) => r[0]?.v);
  assert.deepEqual(labels, [
    'Region',
    'East',
    'East',
    'East Total',
    'West',
    'West Total',
    'Grand Total',
  ]);
});

test('subtotal formulas reference each group by its final absolute rows', () => {
  const out = computeSubtotals({
    rows: SHEET,
    groupCol: 0,
    functionCode: 9,
    subtotalCols: [1],
    startRow: 0,
    startCol: 0,
  });
  // East group is output rows 1..2 → sheet rows 2..3 (1-based) in column B.
  assert.equal(out.rows[3][1]?.f, '=SUBTOTAL(9,B2:B3)');
  // West group is output row 4 → sheet row 5.
  assert.equal(out.rows[5][1]?.f, '=SUBTOTAL(9,B5:B5)');
  // Grand total spans the whole list (header excluded) → B2:B6.
  assert.equal(out.rows[6][1]?.f, '=SUBTOTAL(9,B2:B6)');
});

test('honours the selection origin in formula refs', () => {
  // Selection starting at C3 (startRow=2, startCol=2): group col C, amount col D.
  const rows: SubtotalRow[] = [
    [{ v: 'Region' }, { v: 'Amount' }],
    [{ v: 'East' }, { v: 10 }],
    [{ v: 'West' }, { v: 30 }],
  ];
  const out = computeSubtotals({
    rows,
    groupCol: 0,
    functionCode: 9,
    subtotalCols: [1],
    startRow: 2,
    startCol: 2,
  });
  // East is output row 1 → sheet row 4 (startRow 2 + 1 header + 0). Column D.
  assert.equal(out.rows[2][1]?.f, '=SUBTOTAL(9,D4:D4)');
  // Layout: hdr(3) East(4) EastTotal(5) West(6) WestTotal(7) Grand(8). The grand
  // spans the data + the two subtotal rows (D4:D7); SUBTOTAL ignores the nested ones.
  assert.equal(out.rows[out.rows.length - 1][1]?.f, '=SUBTOTAL(9,D4:D7)');
});

test('uses the chosen function code and labels the group column', () => {
  const out = computeSubtotals({
    rows: SHEET,
    groupCol: 0,
    functionCode: 1, // AVERAGE
    subtotalCols: [1],
    startRow: 0,
    startCol: 0,
  });
  assert.match(out.rows[3][1]?.f ?? '', /^=SUBTOTAL\(1,/);
  assert.equal(out.rows[3][0]?.v, 'East Total');
});

test('empty data is a no-op', () => {
  const out = computeSubtotals({
    rows: [[{ v: 'Region' }, { v: 'Amount' }]],
    groupCol: 0,
    functionCode: 9,
    subtotalCols: [1],
    startRow: 0,
    startCol: 0,
  });
  assert.equal(out.insertedRows, 0);
  assert.equal(out.groups, 0);
});
