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
 * Pivots — Show Values As → % of Grand Total. Each value cell is rewritten as a
 * percentage of that column's grand total. East 350 / West 150 of a 500 total →
 * 70.0% / 30.0%, grand total 100.0%.
 */

async function seed(page: Page) {
  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    const rows = [
      ['Region', 'Product', 'Sales'],
      ['East', 'A', 100],
      ['East', 'B', 200],
      ['East', 'A', 50],
      ['West', 'A', 150],
    ];
    rows.forEach((r, ri) =>
      r.forEach((v, ci) => ws.getRange(ri, ci).setValue({ v: v as string | number })),
    );
    ws.getRange('A1:C5').activate();
  });
}

async function cell(page: Page, a1: string): Promise<unknown> {
  return page.evaluate((c) => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    return ws.getRange(c).getCellData()?.v;
  }, a1);
}

test('Sum of Sales by Region as % of Grand Total', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);
  await seed(page);

  await page.getByTestId('menubar-insert').click();
  await page.getByTestId('menu-item-insert-pivot').click();
  await expect(page.getByTestId('insert-pivot-dialog')).toBeVisible();

  await page.getByTestId('insert-pivot-range').fill('A1:C5');
  await page.getByTestId('insert-pivot-target').fill('E1');
  await page.getByTestId('insert-pivot-row-field').selectOption('0'); // Region
  await page.getByTestId('insert-pivot-value-field').selectOption('2'); // Sales
  await page.getByTestId('insert-pivot-aggregation').selectOption('sum');
  await page.getByTestId('insert-pivot-show-as').selectOption('pctOfGrandTotal');
  await page.getByTestId('insert-pivot-confirm').click();
  await page.waitForTimeout(200);

  expect(await cell(page, 'E1')).toBe('Region');
  expect(await cell(page, 'F1')).toBe('Sum of Sales');
  expect(await cell(page, 'E2')).toBe('East');
  expect(await cell(page, 'F2')).toBe('70.0%');
  expect(await cell(page, 'E3')).toBe('West');
  expect(await cell(page, 'F3')).toBe('30.0%');
  expect(await cell(page, 'E4')).toBe('Grand Total');
  expect(await cell(page, 'F4')).toBe('100.0%');
});
