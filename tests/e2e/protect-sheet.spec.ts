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

import { test, expect } from '@playwright/test';
import { waitForUniver } from './_helpers';
test('Data > Protect blocks edits (read-only)', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);
  await page.evaluate(() =>
    window.__univerAPI!.getActiveWorkbook().getActiveSheet().getRange('A1').setValue({ v: 'before' }));
  await page.waitForTimeout(300);
  await page.getByTestId('menubar-data').click();
  await page.getByTestId('menu-item-protect-sheet').click();
  await page.waitForTimeout(400);
  await page.evaluate(() =>
    window.__univerAPI!.getActiveWorkbook().getActiveSheet().getRange('A1').setValue({ v: 'BLOCKED' }));
  await page.waitForTimeout(500);
  const v = await page.evaluate(() =>
    String(window.__univerAPI!.getActiveWorkbook().getActiveSheet().getRange('A1').getValue() ?? ''));
  console.log('VALUE AFTER PROTECT+EDIT:', v);
  expect(v).toBe('before');
});
