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
  findReference,
  formatValue,
  nextStep,
  stripEquals,
  substitute,
} from './formula-evaluate.js';

test('nextStep finds the innermost-leftmost function call', () => {
  const s = nextStep('SUM(A1:A2)+SQRT(D7)');
  assert.deepEqual(s, { sub: 'SUM(A1:A2)', start: 0, end: 10 });
});

test('nextStep descends into nested calls (innermost first)', () => {
  // ROUND(SUM(A1:A2),2): the inner SUM closes first.
  const s = nextStep('ROUND(SUM(A1:A2),2)');
  assert.equal(s?.sub, 'SUM(A1:A2)');
});

test('nextStep handles a bare grouping paren (no function name)', () => {
  const s = nextStep('(1+2)*3');
  assert.deepEqual(s, { sub: '(1+2)', start: 0, end: 5 });
});

test('nextStep returns null when no parens and no references remain', () => {
  assert.equal(nextStep('3+4*2'), null);
  assert.equal(nextStep('"text"&"more"'), null);
});

test('nextStep steps bare references once parens are gone', () => {
  // Excel underlines each reference before the final arithmetic.
  assert.deepEqual(nextStep('A1+B1'), { sub: 'A1', start: 0, end: 2 });
});

test('findReference finds the leftmost reference, range, and sheet-qualified', () => {
  assert.deepEqual(findReference('A1+B1'), { sub: 'A1', start: 0, end: 2 });
  assert.equal(findReference('10+B2*2')?.sub, 'B2');
  assert.equal(findReference('SUMRESULT+A1:B2')?.sub, 'A1:B2');
  assert.equal(findReference('Sheet2!C3+1')?.sub, 'Sheet2!C3');
  assert.equal(findReference("'My Sheet'!C3")?.sub, "'My Sheet'!C3");
  assert.equal(findReference('$A$1+1')?.sub, '$A$1');
});

test('findReference ignores scientific notation (1E5) and plain numbers', () => {
  assert.equal(findReference('1E5+2'), null); // E5 is an exponent, not a ref
  assert.equal(findReference('3+4*2'), null);
  assert.equal(findReference('2.5e3'), null);
});

test('findReference skips references inside string literals', () => {
  assert.equal(findReference('"A1 is here"&7'), null);
  assert.equal(findReference('"see A1"&B2')?.sub, 'B2');
});

test('nextStep ignores parens inside string literals', () => {
  // The "(x)" is text; the real call is CONCAT(...).
  const s = nextStep('CONCAT("(x)",A1)');
  assert.equal(s?.sub, 'CONCAT("(x)",A1)');
});

test('nextStep respects "" escaped quotes inside strings', () => {
  // The string is `a"b` (escaped), so the ) after it closes IF(...).
  const s = nextStep('IF(A1,"a""b",0)');
  assert.equal(s?.sub, 'IF(A1,"a""b",0)');
});

test('nextStep picks the leftmost of sibling innermost calls', () => {
  const s = nextStep('MAX(A1)+MIN(B1)');
  assert.equal(s?.sub, 'MAX(A1)');
  assert.equal(s?.start, 0);
});

test('substitute replaces the step span', () => {
  const e = 'SUM(A1:A2)+SQRT(D7)';
  const s = nextStep(e)!;
  assert.equal(substitute(e, s.start, s.end, '7'), '7+SQRT(D7)');
});

test('formatValue formats numbers, booleans, strings, errors, empty', () => {
  assert.equal(formatValue(7), '7');
  assert.equal(formatValue(3.5), '3.5');
  assert.equal(formatValue(true), 'TRUE');
  assert.equal(formatValue(false), 'FALSE');
  assert.equal(formatValue('hi'), '"hi"');
  assert.equal(formatValue('a"b'), '"a""b"'); // quote doubled
  assert.equal(formatValue('#DIV/0!'), '#DIV/0!'); // error propagates raw
  assert.equal(formatValue(null), '0');
  assert.equal(formatValue(Infinity), '#NUM!');
});

test('stripEquals removes a leading = and trims', () => {
  assert.equal(stripEquals('=SUM(A1:A2)'), 'SUM(A1:A2)');
  assert.equal(stripEquals('  = A1 + B1 '), 'A1 + B1');
  assert.equal(stripEquals('A1+B1'), 'A1+B1'); // no leading = is fine
});

test('a full stepping sequence reduces with a mock evaluator', () => {
  // Mock: SUM(A1:A2)=7, SQRT(D7)=3, then 7+3=10.
  const evalMock = (sub: string): number => {
    if (sub === 'SUM(A1:A2)') return 7;
    if (sub === 'SQRT(D7)') return 3;
    return 10;
  };
  let e = stripEquals('=SUM(A1:A2)+SQRT(D7)');
  let guard = 0;
  let step = nextStep(e);
  while (step && guard++ < 10) {
    e = substitute(e, step.start, step.end, formatValue(evalMock(step.sub)));
    step = nextStep(e);
  }
  assert.equal(e, '7+3'); // both calls reduced; final arithmetic is one step
  assert.equal(formatValue(evalMock(e)), '10');
});
