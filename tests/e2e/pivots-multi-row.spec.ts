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
 * Pivots P1.5 — multi-row compact layout.
 *
 * Layers a second row field on top of the P0/P1 single-row pivot. The
 * Insert dialog grows one extra dropdown ("Sub-row field"); the
 * resulting pivot uses Excel's compact layout — a single label column
 * with outer keys at the left margin (carrying their subtotal value
 * on the same row) and inner keys indented two spaces (carrying their
 * own aggregated value).
 *
 * Drill-down also has to follow the deeper structure: clicking an
 * outer subtotal row returns every record under that outer key;
 * clicking an indented inner row returns only the records matching
 * the full composite key path.
 */

async function seedRegionProductSales(page: Page) {
  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('A1').setValue({ v: 'Region' });
    ws.getRange('B1').setValue({ v: 'Product' });
    ws.getRange('C1').setValue({ v: 'Sales' });
    const rows: Array<[string, string, number]> = [
      ['East', 'A', 100],
      ['East', 'B', 200],
      ['West', 'A', 150],
      ['East', 'A', 50],   // duplicate East/A so the leaf aggregates
    ];
    rows.forEach(([region, product, sales], i) => {
      ws.getRange(`A${i + 2}`).setValue({ v: region });
      ws.getRange(`B${i + 2}`).setValue({ v: product });
      ws.getRange(`C${i + 2}`).setValue({ v: sales });
    });
  });
}

async function openInsertPivot(page: Page) {
  await page.getByTestId('menubar-insert').click();
  await page.getByTestId('menu-item-insert-pivot').click();
  await expect(page.getByTestId('insert-pivot-dialog')).toBeVisible();
}

async function readCellValue(page: Page, a1: string): Promise<unknown> {
  return page.evaluate((cell) => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    return ws.getRange(cell).getCellData()?.v;
  }, a1);
}

test.describe('Pivots — multi-row compact layout', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await seedRegionProductSales(page);
    await openInsertPivot(page);
  });

  test('Region × Product produces outer subtotals + indented inner rows + grand total', async ({
    page,
  }) => {
    await page.getByTestId('insert-pivot-range').fill('A1:C5');
    await page.getByTestId('insert-pivot-target').fill('E1');
    // Row field = Region (col 0), Sub-row field = Product (col 1),
    // Value field = Sales (col 2), Sum.
    await page.getByTestId('insert-pivot-row-field').selectOption('0');
    await page.getByTestId('insert-pivot-sub-row-field').selectOption('1');
    await page.getByTestId('insert-pivot-value-field').selectOption('2');
    await page.getByTestId('insert-pivot-aggregation').selectOption('sum');
    await page.getByTestId('insert-pivot-confirm').click();
    await page.waitForTimeout(150);

    // Expected compact layout, starting at E1:
    //   E1 = "Region"        F1 = "Sum of Sales"
    //   E2 = "East"          F2 = 350           ← outer subtotal (100+200+50)
    //   E3 = "  A"           F3 = 150           ← inner leaf (100+50)
    //   E4 = "  B"           F4 = 200
    //   E5 = "West"          F5 = 150
    //   E6 = "  A"           F6 = 150
    //   E7 = "Grand Total"   F7 = 500
    expect(await readCellValue(page, 'E1')).toBe('Region');
    expect(await readCellValue(page, 'F1')).toBe('Sum of Sales');
    expect(await readCellValue(page, 'E2')).toBe('East');
    expect(Number(await readCellValue(page, 'F2'))).toBe(350);
    expect(await readCellValue(page, 'E3')).toBe('  A');
    expect(Number(await readCellValue(page, 'F3'))).toBe(150);
    expect(await readCellValue(page, 'E4')).toBe('  B');
    expect(Number(await readCellValue(page, 'F4'))).toBe(200);
    expect(await readCellValue(page, 'E5')).toBe('West');
    expect(Number(await readCellValue(page, 'F5'))).toBe(150);
    expect(await readCellValue(page, 'E6')).toBe('  A');
    expect(Number(await readCellValue(page, 'F6'))).toBe(150);
    expect(await readCellValue(page, 'E7')).toBe('Grand Total');
    expect(Number(await readCellValue(page, 'F7'))).toBe(500);
  });

  test('picking the same column for row + sub-row collapses to single-row', async ({ page }) => {
    // Sanity guard: a user who picks Region in both dropdowns should
    // get a single-row Region pivot, not two stacked Region levels.
    await page.getByTestId('insert-pivot-range').fill('A1:C5');
    await page.getByTestId('insert-pivot-target').fill('E1');
    await page.getByTestId('insert-pivot-row-field').selectOption('0');
    await page.getByTestId('insert-pivot-sub-row-field').selectOption('0');
    await page.getByTestId('insert-pivot-value-field').selectOption('2');
    await page.getByTestId('insert-pivot-confirm').click();
    await page.waitForTimeout(150);

    // No indented rows — same shape as a P0 single-row pivot.
    expect(await readCellValue(page, 'E1')).toBe('Region');
    expect(await readCellValue(page, 'E2')).toBe('East');
    expect(Number(await readCellValue(page, 'F2'))).toBe(350);
    expect(await readCellValue(page, 'E3')).toBe('West');
    expect(Number(await readCellValue(page, 'F3'))).toBe(150);
    expect(await readCellValue(page, 'E4')).toBe('Grand Total');
    expect(Number(await readCellValue(page, 'F4'))).toBe(500);
  });
});
