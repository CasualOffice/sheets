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
import type { IWorkbookData } from '@univerjs/core';

import { readWatchesFromSnapshot, writeWatchesIntoSnapshot } from './watch-resources.js';
import type { Watch } from './watch-model.js';

const WATCH: Watch = { id: 'w1', sheetId: 's1', sheetName: 'Sheet1', row: 0, col: 2 };

function snap(resources: IWorkbookData['resources'] = []): IWorkbookData {
  return { id: 'wb', resources } as IWorkbookData;
}

test('write then read round-trips the watch list', () => {
  const data = snap();
  writeWatchesIntoSnapshot(data, [WATCH]);
  assert.deepEqual(readWatchesFromSnapshot(data), [WATCH]);
});

test('writing an empty list removes the resource', () => {
  const data = snap();
  writeWatchesIntoSnapshot(data, [WATCH]);
  writeWatchesIntoSnapshot(data, []);
  assert.equal(
    data.resources?.find((r) => r.name === '__casual_sheets_watches__'),
    undefined,
  );
  assert.deepEqual(readWatchesFromSnapshot(data), []);
});

test('write replaces a prior watches resource (no duplication)', () => {
  const data = snap();
  writeWatchesIntoSnapshot(data, [WATCH]);
  writeWatchesIntoSnapshot(data, [{ ...WATCH, id: 'w2', row: 5 }]);
  const entries = data.resources?.filter((r) => r.name === '__casual_sheets_watches__') ?? [];
  assert.equal(entries.length, 1);
  assert.deepEqual(
    readWatchesFromSnapshot(data).map((w) => w.id),
    ['w2'],
  );
});

test('read tolerates missing / malformed payloads', () => {
  assert.deepEqual(readWatchesFromSnapshot(undefined), []);
  assert.deepEqual(readWatchesFromSnapshot(snap()), []);
  assert.deepEqual(
    readWatchesFromSnapshot(snap([{ name: '__casual_sheets_watches__', data: 'not json' }])),
    [],
  );
  // Wrong version → ignored.
  assert.deepEqual(
    readWatchesFromSnapshot(
      snap([
        { name: '__casual_sheets_watches__', data: JSON.stringify({ v: 2, watches: [WATCH] }) },
      ]),
    ),
    [],
  );
});

test('read drops malformed individual watches', () => {
  const data = snap([
    {
      name: '__casual_sheets_watches__',
      data: JSON.stringify({ v: 1, watches: [WATCH, { id: 'bad' }] }),
    },
  ]);
  assert.deepEqual(
    readWatchesFromSnapshot(data).map((w) => w.id),
    ['w1'],
  );
});
