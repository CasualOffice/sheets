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
import { selectRange, waitForUniver } from './_helpers';

/**
 * Multi-column sort. Univer's sheets-sort-ui ships a "Custom Sort" panel
 * that supports sort-by-multiple-keys; we surface it from Data → Sort range…
 * which dispatches `sheet.command.sort-range-custom`. We assert on the panel
 * appearing rather than just the command dispatch — sheets-sort-ui routes
 * the call through an internal service and the underlying command id may
 * not always arrive on the public CommandExecuted bus.
 */
test('Data → Sort range… opens the Custom Sort panel', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);

  // Seed a multi-row block so the custom-sort command has a meaningful range
  // (it bails out on a single-cell selection).
  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('A1').setValue({ v: 'Name' });
    ws.getRange('B1').setValue({ v: 'Score' });
    ws.getRange('A2').setValue({ v: 'Charlie' });
    ws.getRange('B2').setValue({ v: 30 });
    ws.getRange('A3').setValue({ v: 'Alice' });
    ws.getRange('B3').setValue({ v: 10 });
    ws.getRange('A4').setValue({ v: 'Bob' });
    ws.getRange('B4').setValue({ v: 20 });
  });
  await selectRange(page, 'A1:B4');

  await page.getByTestId('menubar-data').click();
  await expect(page.getByTestId('menubar-data-popup')).toBeVisible();
  await page.getByTestId('menu-item-sort-custom').click();

  // Univer's panel header / title carries "Sort" wording; one cue is enough
  // to lock down that the panel actually surfaced. Don't depend on the
  // panel's internal markup beyond that — would re-test Univer internals.
  await expect(page.locator('text=/Sort.*Range/i').first()).toBeVisible({
    timeout: 5_000,
  });
});
