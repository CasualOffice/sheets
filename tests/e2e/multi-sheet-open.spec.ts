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
import { waitForUniver } from './_helpers';

/**
 * Regression: opening a multi-sheet xlsx used to surface only the previous
 * workbook's first tab — the new sheets only showed up after the user
 * pressed the "+" Add-sheet button. Root cause: `useSheets` subscribed to
 * sheet-level events only, never to workbook-unit creation, so the swap
 * on File → Open / drag-drop went unnoticed.
 *
 * This test opens a 3-sheet xlsx via the file picker and asserts that
 * all three tabs are visible immediately, before any user interaction.
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

test('Opening a multi-sheet xlsx shows every tab immediately', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);
  await exposeConverters(page);

  // Build a 3-sheet fixture in-page, write it to /tmp, hand to the picker.
  const fixture = '/tmp/casual-sheets-multi-sheet.xlsx';
  const bytes: number[] = await page.evaluate(async () => {
    const snapshot = {
      id: 'wb-multi-1',
      rev: 1,
      name: 'multi',
      appVersion: '0.22.1',
      locale: 1,
      styles: {},
      sheetOrder: ['s-1', 's-2', 's-3'],
      sheets: {
        's-1': { id: 's-1', name: 'Inputs',  cellData: {}, rowCount: 1024, columnCount: 128 },
        's-2': { id: 's-2', name: 'Working', cellData: {}, rowCount: 1024, columnCount: 128 },
        's-3': { id: 's-3', name: 'Output',  cellData: {}, rowCount: 1024, columnCount: 128 },
      },
    };
    const blob = await window.__xlsx!.workbookDataToXlsx(snapshot);
    return Array.from(new Uint8Array(await blob.arrayBuffer()));
  });
  const fs = await import('node:fs');
  fs.writeFileSync(fixture, Buffer.from(bytes));

  await page.getByTestId('menubar-file').click();
  const [chooser] = await Promise.all([
    page.waitForEvent('filechooser'),
    page.getByTestId('menu-item-open').click(),
  ]);
  await chooser.setFiles(fixture);

  // All three tabs must be present without any "+" click prodding.
  // Wait specifically for the third tab so timing flake on the workbook
  // swap doesn't cause a false negative.
  await expect(page.getByText('Output', { exact: true })).toBeVisible({
    timeout: 5_000,
  });
  await expect(page.getByText('Inputs', { exact: true })).toBeVisible();
  await expect(page.getByText('Working', { exact: true })).toBeVisible();
});
