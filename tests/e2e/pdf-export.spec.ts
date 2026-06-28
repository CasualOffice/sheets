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
import { readFileSync } from 'node:fs';
test('File > Download as PDF produces a real PDF', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);
  // put some data so the export isn't empty
  await page.evaluate(() => {
    const api = window.__univerAPI!;
    const ws = api.getActiveWorkbook().getActiveSheet();
    ws.getRange('A1').setValue({ v: 'Item' });
    ws.getRange('B1').setValue({ v: 'Qty' });
    ws.getRange('A2').setValue({ v: 'Widget' });
    ws.getRange('B2').setValue({ v: 42 });
  });
  await page.waitForTimeout(500);
  await page.getByTestId('menubar-file').click();
  const [download] = await Promise.all([
    page.waitForEvent('download', { timeout: 10_000 }),
    page.getByTestId('menu-item-export-pdf').click(),
  ]);
  const name = download.suggestedFilename();
  const path = await download.path();
  const head = readFileSync(path).subarray(0, 5).toString('latin1');
  console.log('PDF DOWNLOAD:', name, ' magic:', head);
  expect(name.endsWith('.pdf')).toBeTruthy();
  expect(head).toBe('%PDF-');
});
