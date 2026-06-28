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

/**
 * Desktop File → Open loads a browser-picked file in-window. A browser File has
 * no real filesystem path, so the window must NOT stay bound to the previously-
 * open file — otherwise the next Save would overwrite that file with the newly-
 * opened content. handleOpen now unbinds the path (Save then prompts).
 */
test('desktop File → Open unbinds the previously-open file path', async ({ page }) => {
  await page.goto('/?desk=1');
  await waitForUniver(page);

  // Simulate a window already bound to a file on disk.
  await page.evaluate(() => {
    (window as unknown as { __deskApp__: { filePath: string | null } }).__deskApp__.filePath =
      '/tmp/original.xlsx';
  });
  expect(
    await page.evaluate(
      () =>
        (window as unknown as { __deskApp__: { filePath: string | null } }).__deskApp__.filePath,
    ),
  ).toBe('/tmp/original.xlsx');

  // Fabricate a fixture .xlsx with a recognizable value.
  await page.evaluate(async () => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const xlsx = await import(/* @vite-ignore */ ('/src/xlsx/index.ts' as any));
    (window as unknown as { __xlsx: unknown }).__xlsx = xlsx;
  });
  const fixtureBytes: number[] = await page.evaluate(async () => {
    const data = {
      id: 'imp',
      rev: 1,
      name: 'imp',
      appVersion: '0.22.1',
      locale: 1,
      styles: {},
      sheetOrder: ['s1'],
      sheets: {
        s1: { id: 's1', name: 'S', cellData: { 0: { 0: { v: 'FILE_X' } } }, rowCount: 1024, columnCount: 128 },
      },
    };
    const blob = await (
      window as unknown as { __xlsx: { workbookDataToXlsx: (d: unknown) => Promise<Blob> } }
    ).__xlsx.workbookDataToXlsx(data);
    return Array.from(new Uint8Array(await blob.arrayBuffer()));
  });
  const fs = await import('node:fs');
  const fixture = '/tmp/casual-sheets-open-unbind.xlsx';
  fs.writeFileSync(fixture, Buffer.from(fixtureBytes));

  // File → Open the fixture in-window.
  await page.getByTestId('menubar-file').click();
  const [chooser] = await Promise.all([
    page.waitForEvent('filechooser'),
    page.getByTestId('menu-item-open').click(),
  ]);
  await chooser.setFiles(fixture);

  // Wait until the picked workbook is actually mounted (parse runs in a worker).
  await page.waitForFunction(
    () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const api = (window as unknown as { __univerAPI?: any }).__univerAPI;
      const ws = api?.getActiveWorkbook()?.getActiveSheet();
      return ws?.getRange('A1').getValue() === 'FILE_X';
    },
    null,
    { timeout: 15_000 },
  );

  // The bound path must now be cleared so a later Save can't overwrite original.xlsx.
  expect(
    await page.evaluate(
      () =>
        (window as unknown as { __deskApp__: { filePath: string | null } }).__deskApp__.filePath,
    ),
  ).toBeNull();
});
