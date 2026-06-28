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

import { expect, test, type Page } from '@playwright/test';
import { createRequire } from 'node:module';
import * as path from 'node:path';
import { waitForUniver } from './_helpers';

const require = createRequire(import.meta.url);
// eslint-disable-next-line @typescript-eslint/no-var-requires
const XLSX = require(require.resolve('@e965/xlsx', { paths: [path.join(process.cwd(), 'apps/web')] }));

declare global {
  interface Window {
    __odsComments?: {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      odsToWorkbookData: (buf: ArrayBuffer) => Promise<any>;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      workbookDataToOds: (data: any) => Promise<Blob>;
    };
  }
}

async function exposeConverters(page: Page) {
  await page.evaluate(async () => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const mod = await import(/* @vite-ignore */ '/src/ods/index.ts' as any);
    window.__odsComments = mod;
  });
}

test.describe('ods comments round-trip', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await exposeConverters(page);
  });

  test('ods authored comments become SHEET_NOTE_PLUGIN resources', async ({ page }) => {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(
      wb,
      {
        A1: { t: 's', v: 'Cell', c: [{ a: 'audit', t: 'ODS comment' }] },
        '!ref': 'A1',
      },
      'Notes',
    );
    const refBytes = Buffer.from(XLSX.write(wb, { type: 'buffer', bookType: 'ods' }));

    const resources = await page.evaluate(async (bytes) => {
      const buf = new Uint8Array(bytes).buffer;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const snap: any = await window.__odsComments!.odsToWorkbookData(buf);
      return snap.resources as Array<{ name: string; data: string }> | undefined;
    }, Array.from(refBytes));

    const notes = resources?.find((r) => r.name === 'SHEET_NOTE_PLUGIN');
    expect(notes, 'SHEET_NOTE_PLUGIN resource present').toBeDefined();
    const parsed = JSON.parse(notes!.data) as Record<string, Record<string, Record<string, { note: string }>>>;
    const firstSheet = Object.values(parsed)[0];
    expect(firstSheet['0']['0'].note).toBe('ODS comment');
  });

  test('snapshot note resources export as ods cell comments', async ({ page }) => {
    const bytes = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const snap: any = {
        id: 'wb-ods-note-1',
        rev: 1,
        name: 'wb',
        appVersion: '0.22.1',
        locale: 1,
        styles: {},
        sheetOrder: ['s1'],
        sheets: {
          s1: {
            id: 's1',
            name: 'Notes',
            cellData: {
              0: {
                0: { v: 'Cell' },
              },
            },
            rowCount: 100,
            columnCount: 26,
          },
        },
        resources: [
          {
            name: 'SHEET_NOTE_PLUGIN',
            data: JSON.stringify({
              s1: {
                0: {
                  0: {
                    id: 'n1',
                    row: 0,
                    col: 0,
                    width: 160,
                    height: 72,
                    note: 'ODS comment',
                  },
                },
              },
            }),
          },
        ],
      };
      const blob = await window.__odsComments!.workbookDataToOds(snap);
      return Array.from(new Uint8Array(await blob.arrayBuffer()));
    });

    const wb = XLSX.read(Buffer.from(bytes), { type: 'buffer' });
    expect(wb.Sheets.Notes.A1?.c?.[0]?.t).toBe('ODS comment');
  });
});
