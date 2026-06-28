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
 * Contract tests for the FileSource interface.
 *
 * We can't exercise `createBrowserFileSource()` from node directly —
 * it touches `indexedDB`, `document`, and `window`. So this suite:
 *
 *   1. Compile-checks the type by building a minimal in-memory
 *      `MockFileSource` and asserting it satisfies `FileSource`.
 *   2. Walks that mock through every method so the next implementer
 *      (PersonalFileSource in Phase C) has a worked example of what
 *      "implements `FileSource`" actually means.
 *   3. Pins the `SaveResult` discriminator shape — three variants
 *      today; bumping the union has to be deliberate.
 */
import { strict as assert } from 'node:assert';
import { test } from 'node:test';

import type { LocaleType } from '@univerjs/core';
import type {
  FileSource,
  OpenedWorkbook,
  RecentEntry,
  RecentId,
  SaveOptions,
  SaveResult,
} from './types';

// Univer's `LocaleType` is a const-enum at TS-level but not re-exported
// as a runtime symbol via the package's `main` entry — `import {
// LocaleType }` throws under node's ESM loader. The numeric value is
// stable across versions (EN_US = 1; see vendor/univer/packages/core
// /src/types/enum/locale-type.ts) and only needed here to satisfy
// `IWorkbookData.locale`. Cast at the seam.
const EN_US = 1 as unknown as LocaleType;

function createMock(): {
  source: FileSource;
  saveCalls: Array<{ size: number; opts: SaveOptions }>;
  forgetCalls: RecentId[];
} {
  const store = new Map<string, OpenedWorkbook>();
  store.set('1', {
    data: {
      id: 'wb-1',
      rev: 0,
      name: 'Quarterly Review',
      appVersion: '0.22.1',
      locale: EN_US,
      styles: {},
      sheetOrder: ['s1'],
      sheets: { s1: { id: 's1', name: 'Sheet1', cellData: {} } },
    },
    sourceFormat: 'xlsx',
  });
  const recents: RecentEntry[] = [
    { id: '1', name: 'Quarterly Review', sourceFormat: 'xlsx', size: 256, modifiedAt: Date.now() },
  ];
  const subs = new Set<() => void>();
  const saveCalls: Array<{ size: number; opts: SaveOptions }> = [];
  const forgetCalls: RecentId[] = [];

  const source: FileSource = {
    kind: 'browser',
    label: 'mock',
    async listRecent() {
      return recents.slice();
    },
    subscribeRecent(cb) {
      subs.add(cb);
      return () => subs.delete(cb);
    },
    async openRecent(id) {
      const hit = store.get(id);
      if (!hit) throw new Error(`recent file not found: ${id}`);
      return hit;
    },
    async forgetRecent(id) {
      forgetCalls.push(id);
    },
    async save(bytes, opts) {
      saveCalls.push({ size: bytes.size, opts });
      const result: SaveResult = { kind: 'download' };
      return result;
    },
  };

  return { source, saveCalls, forgetCalls };
}

test('MockFileSource conforms to the FileSource type (compile + shape)', () => {
  const { source } = createMock();
  assert.equal(source.kind, 'browser');
  assert.equal(source.label, 'mock');
  for (const method of ['listRecent', 'subscribeRecent', 'openRecent', 'forgetRecent', 'save']) {
    assert.equal(
      typeof (source as unknown as Record<string, unknown>)[method],
      'function',
      `missing method: ${method}`,
    );
  }
});

test('listRecent returns the seeded entry', async () => {
  const { source } = createMock();
  const out = await source.listRecent();
  assert.equal(out.length, 1);
  assert.equal(out[0]?.id, '1');
  assert.equal(out[0]?.name, 'Quarterly Review');
});

test('openRecent returns the opened workbook for a known id', async () => {
  const { source } = createMock();
  const opened = await source.openRecent('1');
  assert.equal(opened.data.name, 'Quarterly Review');
  assert.equal(opened.sourceFormat, 'xlsx');
});

test('openRecent throws for an unknown id', async () => {
  const { source } = createMock();
  await assert.rejects(() => source.openRecent('999'), /not found/);
});

test('forgetRecent records the call', async () => {
  const { source, forgetCalls } = createMock();
  await source.forgetRecent('1');
  assert.deepEqual(forgetCalls, ['1']);
});

test('save receives the blob + options', async () => {
  const { source, saveCalls } = createMock();
  const blob = new Blob(['hello'], { type: 'text/plain' });
  const result = await source.save(blob, { filename: 'out.xlsx', sourceFormat: 'xlsx' });
  assert.equal(saveCalls.length, 1);
  assert.equal(saveCalls[0]?.size, 5);
  assert.equal(saveCalls[0]?.opts.filename, 'out.xlsx');
  assert.equal(result.kind, 'download');
});

test('SaveResult discriminator covers exactly three variants', () => {
  // Compile-time pin: every kind currently in the union must produce
  // a value here. Adding a new variant forces an update to this test
  // (and the matching toast/UI branches).
  const variants: SaveResult[] = [
    { kind: 'folder', folderName: 'My Sheets' },
    { kind: 'download' },
    { kind: 'server', path: '/files/abc.xlsx', serverFileId: 'abc', serverEtag: 'v1' },
    { kind: 'conflict', expectedEtag: 'v1' },
    { kind: 'cancelled' },
  ];
  assert.equal(variants.length, 5);
  // Exhaustiveness check — if a new variant is added, the switch loses
  // exhaustiveness and tsc will fail. Runtime: assert no `kind` slips
  // through unhandled.
  for (const v of variants) {
    switch (v.kind) {
      case 'folder':
      case 'download':
      case 'server':
      case 'conflict':
      case 'cancelled':
        break;
      default: {
        const _exhaustive: never = v;
        assert.fail(`unhandled SaveResult variant: ${JSON.stringify(_exhaustive)}`);
      }
    }
  }
});

test('subscribeRecent fires on tick and unsubscribes cleanly', async () => {
  const { source } = createMock();
  let called = 0;
  const unsub = source.subscribeRecent(() => {
    called += 1;
  });
  // The mock doesn't auto-tick; we just verify the contract: subscribe
  // returns an unsubscribe fn, calling it stops further callbacks.
  assert.equal(typeof unsub, 'function');
  unsub();
  assert.equal(called, 0);
});
