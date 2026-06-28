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

import { colToIndex, formulaReferences, parseA1, precedentsOf } from './trace-model.js';

test('colToIndex maps column letters', () => {
  assert.equal(colToIndex('A'), 0);
  assert.equal(colToIndex('Z'), 25);
  assert.equal(colToIndex('AA'), 26);
  assert.equal(colToIndex('AB'), 27);
});

test('parseA1 parses a single cell', () => {
  assert.deepEqual(parseA1('B3'), {
    sheetName: null,
    rect: { startRow: 2, startCol: 1, endRow: 2, endCol: 1 },
  });
});

test('parseA1 parses a range and normalises corners', () => {
  assert.deepEqual(parseA1('C5:A1'), {
    sheetName: null,
    rect: { startRow: 0, startCol: 0, endRow: 4, endCol: 2 },
  });
});

test('parseA1 handles $-absolute and sheet-qualified refs', () => {
  assert.deepEqual(parseA1('$B$3'), {
    sheetName: null,
    rect: { startRow: 2, startCol: 1, endRow: 2, endCol: 1 },
  });
  assert.equal(parseA1('Sheet2!A1')?.sheetName, 'Sheet2');
  assert.equal(parseA1("'My Sheet'!A1")?.sheetName, 'My Sheet');
});

test('parseA1 rejects non-cell shapes', () => {
  assert.equal(parseA1('A:A'), null);
  assert.equal(parseA1('TaxRate'), null);
});

test('precedentsOf returns deduped reference rects', () => {
  const p = precedentsOf('=SUM(A1:A2)+A1*B1');
  // A1:A2, A1 (single), B1 — A1 single and A1:A2 are distinct rects.
  const keys = p.map(
    (t) => `${t.rect.startRow},${t.rect.startCol}-${t.rect.endRow},${t.rect.endCol}`,
  );
  assert.deepEqual(keys.sort(), ['0,0-0,0', '0,0-1,0', '0,1-0,1'].sort());
});

test('precedentsOf ignores numbers and scientific notation', () => {
  assert.deepEqual(precedentsOf('=1E5+2.5'), []);
});

test('formulaReferences detects dependence on a target cell', () => {
  // =A1*2 references A1 on its own sheet.
  assert.equal(formulaReferences('=A1*2', 'Sheet1', 0, 0), true);
  assert.equal(formulaReferences('=A1*2', 'Sheet1', 1, 0), false);
  // A range covers interior cells.
  assert.equal(formulaReferences('=SUM(A1:B2)', 'Sheet1', 1, 1), true);
  // A qualified ref to another sheet doesn't match this sheet's cell.
  assert.equal(formulaReferences('=Sheet2!A1', 'Sheet1', 0, 0), false);
});
