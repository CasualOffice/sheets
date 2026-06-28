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
import { mainCanvas, waitForUniver } from './_helpers';

test('chart sticks to its cell anchor when the grid scrolls', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);

  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('A1').setValue({ v: 'R' });
    ws.getRange('B1').setValue({ v: 'Q1' });
    ws.getRange('A2').setValue({ v: 'N' });
    ws.getRange('B2').setValue({ v: 100 });
    ws.getRange('A3').setValue({ v: 'S' });
    ws.getRange('B3').setValue({ v: 80 });
    ws.getRange('A1:B3').activate();
  });
  await mainCanvas(page).first().click({ position: { x: 100, y: 100 } });
  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('A1:B3').activate();
  });
  await page.getByTestId('menubar-insert').click();
  await page.getByTestId('menu-item-insert-chart').click();
  await page.getByTestId('insert-chart-confirm').click();
  const overlay = page.getByTestId('chart-overlay');
  await expect(overlay).toBeVisible({ timeout: 5_000 });
  const before = await overlay.boundingBox();

  // Scroll a few rows — the chart's anchor cells stay on screen.
  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.scrollToCell(3, 0);
  });
  await page.waitForTimeout(400);

  const after = await overlay.boundingBox();
  expect(before).not.toBeNull();
  expect(after).not.toBeNull();
  // The chart should have moved UP (negative delta) by roughly 3 row
  // heights. Allow a wide range — just prove it moved with the grid
  // instead of staying fixed.
  const dy = after!.y - before!.y;
  expect(dy).toBeLessThan(-30);
});
