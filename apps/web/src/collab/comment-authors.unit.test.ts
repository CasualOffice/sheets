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
import { afterEach, test } from 'node:test';

import {
  __resetCommentAuthors,
  commentAuthorsVersion,
  getCommentAuthor,
  mergeCommentAuthors,
  recordCommentAuthor,
  snapshotCommentAuthors,
  subscribeCommentAuthors,
} from './comment-authors.js';

const A = { name: 'Ada', color: '#1a73e8' };
const B = { name: 'Bo', color: '#d93025' };

afterEach(() => {
  __resetCommentAuthors();
});

test('records and reads an author', () => {
  assert.equal(getCommentAuthor('c1'), undefined);
  assert.equal(recordCommentAuthor('c1', A), true);
  assert.deepEqual(getCommentAuthor('c1'), A);
});

test('an identical re-record is a no-op (no version bump)', () => {
  recordCommentAuthor('c1', A);
  const v = commentAuthorsVersion();
  assert.equal(recordCommentAuthor('c1', { ...A }), false);
  assert.equal(commentAuthorsVersion(), v);
});

test('a real change bumps the version + notifies subscribers', () => {
  let calls = 0;
  const unsub = subscribeCommentAuthors(() => {
    calls += 1;
  });
  const v = commentAuthorsVersion();
  recordCommentAuthor('c1', A);
  assert.equal(commentAuthorsVersion(), v + 1);
  assert.equal(calls, 1);
  unsub();
  recordCommentAuthor('c2', B);
  assert.equal(calls, 1); // unsubscribed — no further notifications
});

test('merge applies a batch and emits once when anything changed', () => {
  let calls = 0;
  subscribeCommentAuthors(() => {
    calls += 1;
  });
  mergeCommentAuthors([
    ['c1', A],
    ['c2', B],
  ]);
  assert.equal(calls, 1);
  assert.deepEqual(getCommentAuthor('c2'), B);
});

test('merge with no changes does not emit', () => {
  recordCommentAuthor('c1', A);
  let calls = 0;
  subscribeCommentAuthors(() => {
    calls += 1;
  });
  mergeCommentAuthors([['c1', { ...A }]]);
  assert.equal(calls, 0);
});

test('snapshot returns every known author for room seeding', () => {
  recordCommentAuthor('c1', A);
  recordCommentAuthor('c2', B);
  assert.deepEqual(
    new Map(snapshotCommentAuthors()),
    new Map([
      ['c1', A],
      ['c2', B],
    ]),
  );
});
