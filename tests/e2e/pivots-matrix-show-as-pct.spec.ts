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
 * Pivots — Show Values As → % of Grand Total in the cross-tab (column-field)
 * layout. Every value cell becomes its share of the single overall grand
 * total. Region × Quarter sums of 500 total → each cell is value/500.
 */

async function cell(page: Page, a1: string): Promise<unknown> {
  return page.evaluate((c) => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    return ws.getRange(c).getCellData()?.v;
  }, a1);
}

test('cross-tab % of Grand Total (Region × Quarter)', async ({ page }) => {
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
  await page.getByTestId('insert-pivot-show-as').selectOption('pctOfGrandTotal');
  await page.getByTestId('insert-pivot-confirm').click();
  await page.waitForTimeout(200);

  // Header: Region | Q1 | Q2 | Grand Total (at E1..H1).
  expect(await cell(page, 'E1')).toBe('Region');
  expect(await cell(page, 'F1')).toBe('Q1');
  expect(await cell(page, 'G1')).toBe('Q2');
  expect(await cell(page, 'H1')).toBe('Grand Total');
  // East: 100/200/300 of 500 → 20% / 40% / 60%.
  expect(await cell(page, 'E2')).toBe('East');
  expect(await cell(page, 'F2')).toBe('20.0%');
  expect(await cell(page, 'G2')).toBe('40.0%');
  expect(await cell(page, 'H2')).toBe('60.0%');
  // West: 150/50/200 → 30% / 10% / 40%.
  expect(await cell(page, 'F3')).toBe('30.0%');
  expect(await cell(page, 'G3')).toBe('10.0%');
  expect(await cell(page, 'H3')).toBe('40.0%');
  // Bottom Grand Total: 250/250/500 → 50% / 50% / 100%.
  expect(await cell(page, 'E4')).toBe('Grand Total');
  expect(await cell(page, 'F4')).toBe('50.0%');
  expect(await cell(page, 'G4')).toBe('50.0%');
  expect(await cell(page, 'H4')).toBe('100.0%');
});
