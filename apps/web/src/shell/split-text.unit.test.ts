import { strict as assert } from 'node:assert';
import { test } from 'node:test';
import {
  DELIMITER,
  buildDelimiterMask,
  hasActiveDelimiter,
  splitPreview,
  type DelimiterOptions,
} from './split-text.ts';

const opts = (o: Partial<DelimiterOptions>): DelimiterOptions => ({
  tab: false,
  comma: false,
  semicolon: false,
  space: false,
  ...o,
});

test('buildDelimiterMask ORs the selected delimiters', () => {
  assert.equal(buildDelimiterMask(opts({ comma: true })), DELIMITER.comma);
  assert.equal(
    buildDelimiterMask(opts({ tab: true, space: true })),
    DELIMITER.tab | DELIMITER.space,
  );
  assert.equal(
    buildDelimiterMask(opts({ comma: true, custom: '|' })),
    DELIMITER.comma | DELIMITER.custom,
  );
  // Empty custom string does not set the custom bit.
  assert.equal(buildDelimiterMask(opts({ comma: true, custom: '' })), DELIMITER.comma);
});

test('hasActiveDelimiter reflects whether anything is selected', () => {
  assert.equal(hasActiveDelimiter(opts({})), false);
  assert.equal(hasActiveDelimiter(opts({ custom: ';' })), true);
});

test('splits on a comma', () => {
  assert.deepEqual(splitPreview(['a,b,c'], opts({ comma: true }), false), [['a', 'b', 'c']]);
});

test('treats consecutive delimiters as one when asked', () => {
  assert.deepEqual(splitPreview(['a,,b'], opts({ comma: true }), false), [['a', '', 'b']]);
  assert.deepEqual(splitPreview(['a,,b'], opts({ comma: true }), true), [['a', 'b']]);
});

test('splits on multiple delimiters at once', () => {
  assert.deepEqual(
    splitPreview(['a,b;c d'], opts({ comma: true, semicolon: true, space: true }), false),
    [['a', 'b', 'c', 'd']],
  );
});

test('custom single-char delimiter', () => {
  assert.deepEqual(splitPreview(['a|b|c'], opts({ custom: '|' }), false), [['a', 'b', 'c']]);
});

test('no active delimiter leaves each row as a single column', () => {
  assert.deepEqual(splitPreview(['a,b'], opts({}), false), [['a,b']]);
});

test('tab delimiter', () => {
  assert.deepEqual(splitPreview(['a\tb'], opts({ tab: true }), false), [['a', 'b']]);
});
