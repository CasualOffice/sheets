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
 * Pure-function tests for the open-file error humanizer. Every entry
 * below is a real-world error string we've actually seen — ExcelJS
 * messages, JSZip messages, fetch failures, gateway HTTP envelopes,
 * or the @casualoffice/core ods loader.
 *
 * Run with: `pnpm test:unit`
 */
import { strict as assert } from 'node:assert';
import { test } from 'node:test';

import { humanizeOpenError } from './humanize-error';

const NAME = 'budget.xlsx';

test('classifies corrupt zip / wrong format', () => {
  const cases = [
    "End of central directory not found",
    "Invalid zip file",
    'central directory record signature missing',
    'Not a zip file',
  ];
  for (const raw of cases) {
    const out = humanizeOpenError(raw, NAME);
    assert.match(out.title, /doesn't look like a valid spreadsheet file/);
    assert.ok(out.hint?.includes('damaged'));
    assert.equal(out.retryLabel, 'Pick a different file');
  }
});

test('classifies encrypted / password protected', () => {
  for (const raw of ['File is encrypted', 'Workbook is password-protected', 'BIFF: unsupported']) {
    const out = humanizeOpenError(raw, NAME);
    assert.match(out.title, /encrypted or in an unsupported format/);
    assert.ok(out.hint?.includes('without the password'));
  }
});

test('classifies network failures', () => {
  for (const raw of ['Failed to fetch', 'NetworkError', 'ECONNREFUSED', 'AbortError: signal']) {
    const out = humanizeOpenError(raw, NAME);
    assert.match(out.title, /Couldn't reach the server/);
    assert.equal(out.retryLabel, 'Try again');
  }
});

test('classifies HTTP 404 as expired-link', () => {
  const out = humanizeOpenError('HTTP 404: not found', NAME);
  assert.match(out.title, /no longer available/);
  assert.ok(out.hint?.includes('share link'));
  assert.equal(out.retryLabel, undefined); // no retry for permanently-gone files
});

test('classifies HTTP 403 as no-access', () => {
  const out = humanizeOpenError('HTTP 403 Forbidden', NAME);
  assert.match(out.title, /don't have access/);
  assert.equal(out.retryLabel, undefined);
});

test('classifies HTTP 5xx as transient', () => {
  for (const raw of ['HTTP 500 internal server error', 'Bad gateway 502']) {
    const out = humanizeOpenError(raw, NAME);
    assert.match(out.title, /Server hiccup/);
    assert.equal(out.retryLabel, 'Try again');
  }
});

test('classifies ods loader errors', () => {
  for (const raw of [
    'expected mimetype application/vnd.oasis.opendocument.spreadsheet',
    'content.xml missing from archive',
  ]) {
    const out = humanizeOpenError(raw, 'budget.ods');
    assert.match(out.title, /doesn't look like a valid \.ods file/);
  }
});

test('classifies memory / size errors', () => {
  for (const raw of ['Out of memory', 'allocation failed', 'RangeError: Invalid array length']) {
    const out = humanizeOpenError(raw, NAME);
    assert.match(out.title, /too large for the browser/);
    assert.equal(out.retryLabel, undefined);
  }
});

test('falls back to friendly generic message', () => {
  const out = humanizeOpenError('Some weird internal failure', NAME);
  assert.match(out.title, /Couldn't open/);
  assert.ok(out.hint && out.hint.length > 0);
  assert.equal(out.retryLabel, 'Try again');
});

test('case-insensitive classification', () => {
  const a = humanizeOpenError('FAILED TO FETCH', NAME);
  const b = humanizeOpenError('failed to fetch', NAME);
  assert.equal(a.title, b.title);
});
