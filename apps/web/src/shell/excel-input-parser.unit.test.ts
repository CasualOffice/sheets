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

/**
 * Pure-function tests for the Excel-style typed-input parser. Verifies
 * the cases the user reported (`$2,000`, `€1.234`) plus the Excel-parity
 * neighbours (percent, accounting parens, suffix locales, edge cases).
 *
 * Run with: `pnpm test:unit`
 */
import { strict as assert } from 'node:assert';
import { test } from 'node:test';

import { parseExcelStyleValue, suggestedPatternFor } from './excel-input-parser';

const parse = parseExcelStyleValue;

test('parses bare dollar currency', () => {
  const r = parse('$2,000');
  assert.equal(r.kind, 'number');
  if (r.kind === 'number') {
    assert.equal(r.value, 2000);
    assert.equal(r.hint, 'currency');
    assert.equal(r.currencySymbol, '$');
  }
});

test('parses dollar with decimal', () => {
  const r = parse('$1,234.56');
  assert.equal(r.kind, 'number');
  if (r.kind === 'number') assert.equal(r.value, 1234.56);
});

test('parses euro prefix', () => {
  const r = parse('€1,500.50');
  assert.equal(r.kind, 'number');
  if (r.kind === 'number') {
    assert.equal(r.value, 1500.5);
    assert.equal(r.currencySymbol, '€');
  }
});

test('parses pound prefix', () => {
  const r = parse('£99.99');
  assert.equal(r.kind, 'number');
  if (r.kind === 'number') assert.equal(r.value, 99.99);
});

test('parses yen / yuan prefix (no decimal)', () => {
  const r = parse('¥10,000');
  assert.equal(r.kind, 'number');
  if (r.kind === 'number') assert.equal(r.value, 10000);
});

test('parses rupee prefix', () => {
  const r = parse('₹50,000');
  assert.equal(r.kind, 'number');
  if (r.kind === 'number') assert.equal(r.value, 50000);
});

test('parses ruble prefix', () => {
  const r = parse('₽1,200');
  assert.equal(r.kind, 'number');
  if (r.kind === 'number') assert.equal(r.value, 1200);
});

test('parses won prefix', () => {
  const r = parse('₩500');
  assert.equal(r.kind, 'number');
  if (r.kind === 'number') assert.equal(r.value, 500);
});

test('parses NT$ multi-char prefix without consuming as $ only', () => {
  const r = parse('NT$1,234');
  assert.equal(r.kind, 'number');
  if (r.kind === 'number') {
    assert.equal(r.value, 1234);
    assert.equal(r.currencySymbol, 'NT$');
  }
});

test('parses R$ multi-char prefix (Brazil real)', () => {
  const r = parse('R$1,500');
  assert.equal(r.kind, 'number');
  if (r.kind === 'number') {
    assert.equal(r.value, 1500);
    assert.equal(r.currencySymbol, 'R$');
  }
});

test('parses CHF prefix (Swiss franc)', () => {
  const r = parse('CHF250');
  assert.equal(r.kind, 'number');
  if (r.kind === 'number') assert.equal(r.value, 250);
});

test('parses kr suffix (Nordic)', () => {
  const r = parse('100 kr');
  assert.equal(r.kind, 'number');
  if (r.kind === 'number') {
    assert.equal(r.value, 100);
    assert.equal(r.currencySymbol, 'kr');
  }
});

test('parses zł suffix (Polish złoty)', () => {
  const r = parse('250 zł');
  assert.equal(r.kind, 'number');
  if (r.kind === 'number') assert.equal(r.value, 250);
});

test('parses €-suffix locale shape', () => {
  const r = parse('1,500 €');
  assert.equal(r.kind, 'number');
  if (r.kind === 'number') assert.equal(r.value, 1500);
});

test('parses trailing percent', () => {
  const r = parse('15%');
  assert.equal(r.kind, 'number');
  if (r.kind === 'number') {
    assert.equal(r.value, 0.15);
    assert.equal(r.hint, 'percent');
  }
});

test('parses fractional percent', () => {
  const r = parse('12.5%');
  assert.equal(r.kind, 'number');
  if (r.kind === 'number') assert.equal(r.value, 0.125);
});

test('parses negative percent', () => {
  const r = parse('-3%');
  assert.equal(r.kind, 'number');
  if (r.kind === 'number') assert.equal(r.value, -0.03);
});

test('parses accounting parentheses as negative', () => {
  const r = parse('(500)');
  assert.equal(r.kind, 'number');
  if (r.kind === 'number') {
    assert.equal(r.value, -500);
    assert.equal(r.hint, 'accounting');
  }
});

test('parses parens with inner currency', () => {
  const r = parse('($1,234.50)');
  assert.equal(r.kind, 'number');
  if (r.kind === 'number') {
    assert.equal(r.value, -1234.5);
    assert.equal(r.hint, 'currency-accounting');
    assert.equal(r.currencySymbol, '$');
  }
});

test('parses leading plus sign', () => {
  const r = parse('+500');
  assert.equal(r.kind, 'number');
  if (r.kind === 'number') assert.equal(r.value, 500);
});

test('parses negative dollar', () => {
  const r = parse('-$1,234');
  assert.equal(r.kind, 'number');
  if (r.kind === 'number') assert.equal(r.value, -1234);
});

test('parses dollar with sign INSIDE the symbol (Excel accepts this)', () => {
  const r = parse('$-1,234');
  assert.equal(r.kind, 'number');
  if (r.kind === 'number') assert.equal(r.value, -1234);
});

test('parses thousands grouping without symbol', () => {
  const r = parse('1,234,567');
  assert.equal(r.kind, 'number');
  if (r.kind === 'number') {
    assert.equal(r.value, 1234567);
    assert.equal(r.hint, 'plain');
  }
});

test('parses scientific notation (passes through Number())', () => {
  const r = parse('1.5e3');
  assert.equal(r.kind, 'number');
  if (r.kind === 'number') assert.equal(r.value, 1500);
});

test('leaves invalid thousands grouping as text (1,5 — not a number in en-US)', () => {
  const r = parse('1,5');
  assert.equal(r.kind, 'unchanged');
});

test('leaves bare currency symbol as text', () => {
  const r = parse('$');
  assert.equal(r.kind, 'unchanged');
});

test('leaves apostrophe-prefixed input alone (force-text marker)', () => {
  const r = parse("'42");
  assert.equal(r.kind, 'unchanged');
});

test('leaves formula input alone', () => {
  const r = parse('=SUM(A1:A10)');
  assert.equal(r.kind, 'unchanged');
});

test('leaves plain text alone', () => {
  const r = parse('Hello world');
  assert.equal(r.kind, 'unchanged');
});

test('leaves empty / whitespace-only input alone', () => {
  assert.equal(parse('').kind, 'unchanged');
  assert.equal(parse('   ').kind, 'unchanged');
});

test('does not turn alphanumerics into numbers (e.g. ZIP-like)', () => {
  assert.equal(parse('12K').kind, 'unchanged');
  assert.equal(parse('A1').kind, 'unchanged');
});

test('preserves negative zero as zero (not -0 quirk leaking)', () => {
  const r = parse('-0');
  assert.equal(r.kind, 'number');
  if (r.kind === 'number') assert.equal(r.value, 0);
});

test('suggestedPatternFor returns currency pattern with symbol', () => {
  assert.equal(suggestedPatternFor('currency', '€'), '"€"#,##0.00');
  assert.equal(suggestedPatternFor('currency', '$'), '"$"#,##0.00');
});

test('suggestedPatternFor returns percent pattern', () => {
  assert.equal(suggestedPatternFor('percent', undefined), '0.00%');
});

test('suggestedPatternFor returns accounting bracket pattern', () => {
  assert.equal(suggestedPatternFor('accounting', undefined), '#,##0.00_);(#,##0.00)');
});

test('suggestedPatternFor combines currency + accounting', () => {
  const p = suggestedPatternFor('currency-accounting', '£');
  assert.ok(p?.includes('"£"'));
  assert.ok(p?.includes('('));
});

test('suggestedPatternFor returns null for plain (no pattern needed)', () => {
  assert.equal(suggestedPatternFor('plain', undefined), null);
});
