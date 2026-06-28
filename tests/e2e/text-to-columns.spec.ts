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
 * Text to Columns (Data → Text to Columns…). Splits a single column of
 * delimited text into several columns, with a live preview. The split runs
 * through the Univer fork's split-text-to-columns command (inserts columns,
 * undoable); delimiter parsing is unit-tested in split-text.ts. This drives the
 * wizard end-to-end.
 */

async function seedColumn(page: import('@playwright/test').Page, values: string[]) {
  await page.evaluate((vals) => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    vals.forEach((v, r) => ws.getRange(r, 0).setValue({ v }));
    ws.getRange(0, 0, vals.length, 1).activate();
  }, values);
}

async function rowValues(page: import('@playwright/test').Page, row: number, cols: number) {
  return page.evaluate(
    ({ row, cols }) => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      const out: Array<unknown> = [];
      for (let c = 0; c < cols; c++) out.push(ws.getRange(row, c).getCellData()?.v ?? null);
      return out;
    },
    { row, cols },
  );
}

test('splits a comma-delimited column into separate columns', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);
  await seedColumn(page, ['Ada,Lovelace,1815', 'Bob,Stone,1990']);

  await page.getByTestId('menubar-data').click();
  await page.getByTestId('menu-item-text-to-columns').click();
  await expect(page.getByTestId('text-to-columns-dialog')).toBeVisible();

  // Comma is on by default; the preview shows three columns.
  await expect(page.getByTestId('ttc-preview')).toContainText('Lovelace');
  await page.getByTestId('ttc-finish').click();
  await expect(page.getByTestId('text-to-columns-dialog')).toBeHidden();

  expect(await rowValues(page, 0, 3)).toEqual(['Ada', 'Lovelace', 1815]);
  expect(await rowValues(page, 1, 3)).toEqual(['Bob', 'Stone', 1990]);
});

test('splits on a custom delimiter', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);
  await seedColumn(page, ['x|y|z']);

  await page.getByTestId('menubar-data').click();
  await page.getByTestId('menu-item-text-to-columns').click();
  // Turn comma off, enable Other = '|'.
  await page.getByTestId('ttc-comma').uncheck();
  await page.getByTestId('ttc-other').check();
  await page.getByTestId('ttc-other-input').fill('|');
  await page.getByTestId('ttc-finish').click();

  expect(await rowValues(page, 0, 3)).toEqual(['x', 'y', 'z']);
});

test('warns when more than one column is selected', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);
  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('A1:B2').activate();
  });

  await page.getByTestId('menubar-data').click();
  await page.getByTestId('menu-item-text-to-columns').click();
  await expect(page.getByTestId('ttc-single-col-notice')).toBeVisible();
  await expect(page.getByTestId('ttc-finish')).toBeDisabled();
});
