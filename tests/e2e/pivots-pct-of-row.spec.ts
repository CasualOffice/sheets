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
 * Pivots — Show Values As → % of Row Total (cross-tab). Each value cell becomes
 * its share of its row's total, so every row sums to 100% across the column
 * keys. Region × Quarter: East row = 300, West row = 200.
 */

async function cell(page: Page, a1: string): Promise<unknown> {
  return page.evaluate((c) => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    return ws.getRange(c).getCellData()?.v;
  }, a1);
}

test('cross-tab % of Row Total — rows each sum to 100%', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);
  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    const rows = [
      ['Region', 'Quarter', 'Sales'],
      ['East', 'Q1', 100],
      ['East', 'Q2', 200],
      ['West', 'Q1', 150],
      ['West', 'Q2', 50],
    ];
    rows.forEach((r, ri) =>
      r.forEach((v, ci) => ws.getRange(ri, ci).setValue({ v: v as string | number })),
    );
    ws.getRange('A1:C5').activate();
  });

  await page.getByTestId('menubar-insert').click();
  await page.getByTestId('menu-item-insert-pivot').click();
  await expect(page.getByTestId('insert-pivot-dialog')).toBeVisible();

  await page.getByTestId('insert-pivot-range').fill('A1:C5');
  await page.getByTestId('insert-pivot-target').fill('E1');
  await page.getByTestId('insert-pivot-row-field').selectOption('0'); // Region
  await page.getByTestId('insert-pivot-col-field').selectOption('1'); // Quarter
  await page.getByTestId('insert-pivot-value-field').selectOption('2'); // Sales
  await page.getByTestId('insert-pivot-aggregation').selectOption('sum');
  await page.getByTestId('insert-pivot-show-as').selectOption('pctOfRowTotal');
  await page.getByTestId('insert-pivot-confirm').click();
  await page.waitForTimeout(200);

  // East row total 300: 100/300=33.3%, 200/300=66.7%, grand 100%.
  expect(await cell(page, 'F2')).toBe('33.3%');
  expect(await cell(page, 'G2')).toBe('66.7%');
  expect(await cell(page, 'H2')).toBe('100.0%');
  // West row total 200: 150/200=75%, 50/200=25%, grand 100%.
  expect(await cell(page, 'F3')).toBe('75.0%');
  expect(await cell(page, 'G3')).toBe('25.0%');
  expect(await cell(page, 'H3')).toBe('100.0%');
  // Grand Total row total 500: 250/500=50%, 250/500=50%, grand 100%.
  expect(await cell(page, 'F4')).toBe('50.0%');
  expect(await cell(page, 'G4')).toBe('50.0%');
  expect(await cell(page, 'H4')).toBe('100.0%');
});
