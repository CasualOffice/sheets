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
 * Pivots — date grouping. A date row field can be grouped by Year / Quarter /
 * Month; records bucket by the derived period instead of the raw date. Dates
 * are seeded via =DATE() so the cells hold real Excel serials.
 */

async function cell(page: Page, a1: string): Promise<unknown> {
  return page.evaluate((c) => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    return ws.getRange(c).getCellData()?.v;
  }, a1);
}

test('grouping a date row field by Years buckets records per year', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);

  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('A1').setValue({ v: 'Date' });
    ws.getRange('B1').setValue({ v: 'Sales' });
    ws.getRange('A2').setValue({ f: '=DATE(2025,1,10)' });
    ws.getRange('A3').setValue({ f: '=DATE(2025,6,20)' });
    ws.getRange('A4').setValue({ f: '=DATE(2026,3,5)' });
    ws.getRange('B2').setValue({ v: 100 });
    ws.getRange('B3').setValue({ v: 200 });
    ws.getRange('B4').setValue({ v: 50 });
  });

  // Wait for the DATE() formulas to resolve (they surface as a date string
  // like "2025/01/10" in this app).
  await page.waitForFunction(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    return /^2025/.test(String(ws.getRange('A2').getValue() ?? ''));
  });

  await page.getByTestId('menubar-insert').click();
  await page.getByTestId('menu-item-insert-pivot').click();
  await expect(page.getByTestId('insert-pivot-dialog')).toBeVisible();

  await page.getByTestId('insert-pivot-range').fill('A1:B4');
  await page.getByTestId('insert-pivot-target').fill('D1');
  await page.getByTestId('insert-pivot-row-field').selectOption('0'); // Date
  await page.getByTestId('insert-pivot-row-grouping').selectOption('year');
  await page.getByTestId('insert-pivot-value-field').selectOption('1'); // Sales
  await page.getByTestId('insert-pivot-aggregation').selectOption('sum');
  await page.getByTestId('insert-pivot-confirm').click();
  await page.waitForTimeout(200);

  // Grouped by year: 2025 = 100+200 = 300, 2026 = 50, Grand = 350.
  expect(await cell(page, 'D1')).toBe('Date');
  expect(await cell(page, 'E1')).toBe('Sum of Sales');
  // Year keys are numeric strings; the cell write coerces them to numbers.
  expect(Number(await cell(page, 'D2'))).toBe(2025);
  expect(Number(await cell(page, 'E2'))).toBe(300);
  expect(Number(await cell(page, 'D3'))).toBe(2026);
  expect(Number(await cell(page, 'E3'))).toBe(50);
  expect(await cell(page, 'D4')).toBe('Grand Total');
  expect(Number(await cell(page, 'E4'))).toBe(350);
});
