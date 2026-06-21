/**
 * Pure-function tests for File → Properties helpers. These cover the
 * regressions reported against the Properties dialog:
 *   - placeholder junk ("Unknown" creator, literal "null"/"undefined")
 *     leaking into metadata fields,
 *   - malformed timestamps rendering as the literal "Invalid Date",
 *   - byte-size formatting (the dialog now shows the real uploaded size,
 *     not the much-larger uncompressed JSON snapshot).
 *
 * Run with: `pnpm test:unit`
 */
import { strict as assert } from 'node:assert';
import { test } from 'node:test';

import { formatBytes, formatDate, sanitizeProps } from './file-menu';

test('sanitizeProps drops placeholder + blank values', () => {
  const out = sanitizeProps({
    title: '',
    author: 'Unknown',
    subject: '  ',
    category: 'null',
    company: 'undefined',
    tags: '  Q3, budget ',
    description: 'Real description',
  });
  // Placeholders and blanks are gone entirely (not rendered as empty strings).
  assert.deepEqual(Object.keys(out).sort(), ['description', 'tags']);
  // Surviving values are trimmed.
  assert.equal(out.tags, 'Q3, budget');
  assert.equal(out.description, 'Real description');
});

test('sanitizeProps treats placeholders case-insensitively', () => {
  const out = sanitizeProps({ author: 'UNKNOWN', title: 'Null' });
  assert.deepEqual(out, {});
});

test('formatDate guards against Invalid Date', () => {
  assert.equal(formatDate(undefined), '—');
  assert.equal(formatDate(''), '—');
  assert.equal(formatDate('not-a-date'), '—');
  // A valid ISO string renders to *something* locale-specific, never the
  // literal "Invalid Date".
  const ok = formatDate('2026-06-20T12:00:00.000Z');
  assert.notEqual(ok, '—');
  assert.doesNotMatch(ok, /Invalid Date/);
});

test('formatBytes scales B / KB / MB', () => {
  assert.equal(formatBytes(512), '512 B');
  assert.equal(formatBytes(2048), '2.0 KB');
  assert.equal(formatBytes(6 * 1024 * 1024), '6.00 MB');
});
