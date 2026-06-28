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
test('Format Cells > Text orientation applies rotation', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);
  await page.evaluate(() => {
    const ws = window.__univerAPI!.getActiveWorkbook().getActiveSheet();
    ws.getRange('A1').setValue({ v: 'rotated' });
    ws.getRange('A1').activate();
  });
  await page.getByTestId('menubar-format').click();
  await page.getByTestId('menu-item-format-cells').click();
  await page.getByTestId('format-cells-dialog').waitFor({ timeout: 8000 });
  await page.getByTestId('format-cells-tab-alignment').click();
  await page.getByTestId('format-cells-rotation').selectOption('45');
  await page.getByTestId('format-cells-apply').click();
  await page.waitForTimeout(600);
  const angle = await page.evaluate(() => {
    const api = window.__univerAPI!;
    const ws = api.getActiveWorkbook().getActiveSheet();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const cd: any = ws.getRange(0, 0).getCellData?.();
    const s = cd?.s;
    const style =
      typeof s === 'string' ? api.getActiveWorkbook().getWorkbook().getStyles().get(s) : s;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return (style as any)?.tr?.a ?? null;
  });
  console.log('ROTATION ANGLE:', angle);
  expect(angle).toBe(45);
});
