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
 * Pivot drill-down — Ctrl+Shift+D on a pivot result cell pops the
 * source rows that fed it. v1 dispatches against the active selection
 * (no canvas double-click hook needed).
 */

async function seedAndPivot(page: Page) {
  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    // Headers
    ws.getRange('A1').setValue({ v: 'Region' });
    ws.getRange('B1').setValue({ v: 'Sales' });
    // Records
    ws.getRange('A2').setValue({ v: 'North' });
    ws.getRange('B2').setValue({ v: 10 });
    ws.getRange('A3').setValue({ v: 'South' });
    ws.getRange('B3').setValue({ v: 20 });
    ws.getRange('A4').setValue({ v: 'North' });
    ws.getRange('B4').setValue({ v: 30 });
  });
  await page.getByTestId('menubar-insert').click();
  await page.getByTestId('menu-item-insert-pivot').click();
  await page.getByTestId('insert-pivot-range').fill('A1:B4');
  await page.getByTestId('insert-pivot-target').fill('D1');
  await page.getByTestId('insert-pivot-row-field').selectOption('0');
  await page.getByTestId('insert-pivot-value-field').selectOption('1');
  await page.getByTestId('insert-pivot-aggregation').selectOption('sum');
  await page.getByTestId('insert-pivot-confirm').click();
  await page.waitForTimeout(150);
}

test.describe('Pivots — drill-down', () => {
  test('drilling on a row key pops the contributing source records', async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await seedAndPivot(page);

    // Pivot layout starting at D1:
    //   D1: "Region",   E1: "Sum of Sales"
    //   D2: "North",    E2: 40
    //   D3: "South",    E3: 20
    //   D4: "Grand Total", E4: 60
    // Select D2 (North row), drill down.
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('D2').activate();
    });
    await page.keyboard.press('Control+Shift+D');
    await expect(page.getByTestId('drill-down-dialog')).toBeVisible({ timeout: 3_000 });
    await expect(page.getByTestId('drill-down-dialog')).toContainText('Region = "North"');
    // Two North records: 10 + 30.
    const rows = page.getByTestId('drill-down-row');
    await expect(rows).toHaveCount(2);
    const text = await page.getByTestId('drill-down-body').innerText();
    expect(text).toContain('10');
    expect(text).toContain('30');
    expect(text).not.toContain('20'); // South shouldn't be here
  });

  test('drilling on Grand Total returns every filtered record', async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await seedAndPivot(page);
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('D4').activate(); // Grand Total row
    });
    await page.keyboard.press('Control+Shift+D');
    await expect(page.getByTestId('drill-down-dialog')).toBeVisible({ timeout: 3_000 });
    await expect(page.getByTestId('drill-down-dialog')).toContainText('Grand Total · 3 rows');
    await expect(page.getByTestId('drill-down-row')).toHaveCount(3);
  });

  test('drilling outside a pivot is a silent no-op', async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await seedAndPivot(page);
    // Select a cell that's outside the pivot output (e.g. G10).
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('G10').activate();
    });
    await page.keyboard.press('Control+Shift+D');
    await page.waitForTimeout(200);
    await expect(page.getByTestId('drill-down-dialog')).toHaveCount(0);
  });
});
