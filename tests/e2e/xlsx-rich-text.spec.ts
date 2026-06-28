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

import { expect, test } from '@playwright/test';
import { waitForUniver } from './_helpers';
import type { Page } from '@playwright/test';

/**
 * In-cell rich text xlsx round-trip — Phase 4, T4.2.
 *
 * A cell with mixed per-character formatting (a bold word inside otherwise
 * plain text) is stored as `cell.p` (an IDocumentBody with textRuns). The SDK
 * used to drop this on export and flatten it to a plain string on import; now
 * it maps to/from ExcelJS `richText` so the formatting survives a round-trip
 * (rich-text.ts). This guards the "lossless xlsx round-trip" moat.
 */

declare global {
  interface Window {
    __xlsx?: {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      xlsxToWorkbookData: (buf: ArrayBuffer) => Promise<any>;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      workbookDataToXlsx: (data: any) => Promise<Blob>;
    };
  }
}

async function exposeConverters(page: Page) {
  await page.evaluate(async () => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const mod = await import(/* @vite-ignore */ '/src/xlsx/index.ts' as any);
    window.__xlsx = mod;
  });
}

test('a cell with a bold word survives export → re-import', async ({ page }) => {
  test.setTimeout(60_000);
  await page.goto('/');
  await waitForUniver(page);
  await exposeConverters(page);

  const reloaded = await page.evaluate(async () => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    // "Hi" bold + "there" normal in A1.
    ws.getRange('A1').setValue({
      v: 'Hithere',
      p: {
        id: '__INTERNAL_EDITOR__DOCS_NORMAL',
        documentStyle: {},
        body: {
          dataStream: 'Hithere\r\n',
          textRuns: [
            { st: 0, ed: 2, ts: { bl: 1 } },
            { st: 2, ed: 7 },
          ],
          paragraphs: [{ startIndex: 7 }],
          sectionBreaks: [{ startIndex: 8 }],
        },
      },
    });
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const original: any = api.getActiveWorkbook()!.save();
    const blob = await window.__xlsx!.workbookDataToXlsx(original);
    const buf = await blob.arrayBuffer();
    return window.__xlsx!.xlsxToWorkbookData(buf);
  });

  const sheetId = reloaded.sheetOrder[0];
  const a1 = reloaded.sheets[sheetId].cellData['0']['0'];

  // Text preserved, and a rich-text body came back with a bold first run.
  expect(a1.v).toBe('Hithere');
  expect(a1.p?.body?.dataStream).toBe('Hithere\r\n');
  const runs = a1.p?.body?.textRuns ?? [];
  const boldRun = runs.find((r: { ts?: { bl?: number } }) => r.ts?.bl === 1);
  expect(boldRun).toBeTruthy();
  expect(boldRun.st).toBe(0);
  expect(boldRun.ed).toBe(2);
});

test('a plain cell is NOT promoted to rich text (no spurious p body)', async ({ page }) => {
  test.setTimeout(60_000);
  await page.goto('/');
  await waitForUniver(page);
  await exposeConverters(page);

  const reloaded = await page.evaluate(async () => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('A1').setValue({ v: 'plain' });
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const original: any = api.getActiveWorkbook()!.save();
    const blob = await window.__xlsx!.workbookDataToXlsx(original);
    const buf = await blob.arrayBuffer();
    return window.__xlsx!.xlsxToWorkbookData(buf);
  });

  const sheetId = reloaded.sheetOrder[0];
  const a1 = reloaded.sheets[sheetId].cellData['0']['0'];
  expect(a1.v).toBe('plain');
  expect(a1.p).toBeFalsy(); // stayed a plain value
});
