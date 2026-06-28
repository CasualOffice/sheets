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
 * Sparklines — in-cell mini-charts. Verifies the insert dialog wiring
 * and that the SparklineLayer renders an SVG for the new model.
 */

test.describe('Sparklines', () => {
  test('Insert → Sparkline creates an in-cell SVG', async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    // Seed a row of values.
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('A1').setValue({ v: 1 });
      ws.getRange('B1').setValue({ v: 5 });
      ws.getRange('C1').setValue({ v: 3 });
      ws.getRange('D1').setValue({ v: 8 });
      ws.getRange('E1').setValue({ v: 2 });
    });
    await page.getByTestId('menubar-insert').click();
    await page.getByTestId('menu-item-insert-sparkline').click();
    await expect(page.getByTestId('insert-sparkline-dialog')).toBeVisible();
    await page.getByTestId('insert-sparkline-source').fill('A1:E1');
    await page.getByTestId('insert-sparkline-anchor').fill('G1');
    await page.getByTestId('insert-sparkline-ok').click();
    // Layer + cell should render shortly.
    await expect(page.getByTestId('sparkline-layer')).toBeVisible({ timeout: 5_000 });
    await expect(page.getByTestId('sparkline-cell')).toHaveCount(1);
  });

  test('Insert rejects a bad range', async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await page.getByTestId('menubar-insert').click();
    await page.getByTestId('menu-item-insert-sparkline').click();
    await page.getByTestId('insert-sparkline-source').fill('xyz');
    await page.getByTestId('insert-sparkline-anchor').fill('G1');
    await page.getByTestId('insert-sparkline-ok').click();
    await expect(page.getByTestId('insert-sparkline-error')).toBeVisible();
  });
});
