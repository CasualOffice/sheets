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
import { compileCriterion, matchAdvancedFilter } from './advanced-filter.ts';

test('comparison operators on numbers', () => {
  assert.equal(compileCriterion('>100')(150), true);
  assert.equal(compileCriterion('>100')(50), false);
  assert.equal(compileCriterion('>=100')(100), true);
  assert.equal(compileCriterion('<=5')(5), true);
  assert.equal(compileCriterion('<5')(5), false);
});

test('equality, inequality and exact match', () => {
  assert.equal(compileCriterion('=East')('East'), true);
  assert.equal(compileCriterion('=East')('Eastern'), false); // end-anchored
  assert.equal(compileCriterion('<>East')('West'), true);
  assert.equal(compileCriterion('<>East')('East'), false);
});

test('bare text is begins-with; bare number is equality', () => {
  assert.equal(compileCriterion('Ea')('Eastern'), true); // begins with
  assert.equal(compileCriterion('Ea')('West'), false);
  assert.equal(compileCriterion('90')(90), true);
  assert.equal(compileCriterion('90')(91), false);
});

test('wildcards', () => {
  assert.equal(compileCriterion('=A*n')('Aileen'), true);
  assert.equal(compileCriterion('=A?c')('Abc'), true);
  assert.equal(compileCriterion('=A?c')('Ac'), false);
  assert.equal(compileCriterion('=10~%')('10%'), true); // ~ escapes the literal %
});

test('empty criterion matches everything', () => {
  assert.equal(compileCriterion('')('anything'), true);
  assert.equal(compileCriterion('  ')(42), true);
});

// list: Region, Amount
const LIST_HEADER = ['Region', 'Amount'];
const LIST_ROWS = [
  ['East', 90],
  ['West', 40],
  ['East', 20],
  ['North', 200],
];

test('AND across columns within a criteria row', () => {
  // Region = East AND Amount > 50
  const idx = matchAdvancedFilter({
    listHeader: LIST_HEADER,
    listRows: LIST_ROWS,
    critHeader: ['Region', 'Amount'],
    critRows: [['East', '>50']],
  });
  assert.deepEqual(idx, [0]); // only East/90
});

test('OR across criteria rows', () => {
  // Region = West  OR  Amount >= 200
  const idx = matchAdvancedFilter({
    listHeader: LIST_HEADER,
    listRows: LIST_ROWS,
    critHeader: ['Region', 'Amount'],
    critRows: [
      ['=West', ''],
      ['', '>=200'],
    ],
  });
  assert.deepEqual(idx, [1, 3]); // West/40 and North/200
});

test('criteria headers are matched by name, not position', () => {
  // Criteria range only has the Amount column.
  const idx = matchAdvancedFilter({
    listHeader: LIST_HEADER,
    listRows: LIST_ROWS,
    critHeader: ['Amount'],
    critRows: [['>50']],
  });
  assert.deepEqual(idx, [0, 3]); // 90 and 200
});

test('unknown criteria columns are ignored', () => {
  const idx = matchAdvancedFilter({
    listHeader: LIST_HEADER,
    listRows: LIST_ROWS,
    critHeader: ['Nonsense'],
    critRows: [['x']],
  });
  // The only condition references a missing column → that criteria row has no
  // usable conditions → matches everything.
  assert.deepEqual(idx, [0, 1, 2, 3]);
});
