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
 * Data → Show all rows. Recovery action for the case where a filter (sheet- or
 * table-level) left rows hidden after the filter UI was dismissed and there's
 * no obvious surface to unhide them from.
 */
test('Data → Show all rows reveals every hidden row', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);

  await selectRange(page, 'A1');
  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    for (let r = 0; r < 6; r++) ws.getRange(r, 0).setValue({ v: `r${r + 1}` });
    // Hide rows 2..4 (0-indexed: 1..3).
    ws.hideRows(1, 3);
  });

  // FWorksheet doesn't expose isRowHidden; reach through to the underlying
  // Worksheet, which has getRowVisible(row).
  const hiddenBefore = await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const fws: any = api.getActiveWorkbook()!.getActiveSheet();
    const ws = fws.getSheet();
    return [1, 2, 3].map((r) => !ws.getRowVisible(r));
  });
  expect(hiddenBefore).toEqual([true, true, true]);

  await page.getByTestId('menubar-data').click();
  await page.getByTestId('menu-item-show-all-rows').click();

  const hiddenAfter = await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const fws: any = api.getActiveWorkbook()!.getActiveSheet();
    const ws = fws.getSheet();
    return [1, 2, 3].map((r) => !ws.getRowVisible(r));
  });
  expect(hiddenAfter).toEqual([false, false, false]);
});
