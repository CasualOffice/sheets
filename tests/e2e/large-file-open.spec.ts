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
 * Large-file open validation — Phase 1, T1.4 (docs/COMPETITIVE_ROADMAP.md).
 *
 * Opening a file with far more rows than the initial grid extent must load the
 * full sheet (the import sets `rowCount` to the true data extent, un-clamped —
 * `parse-impl.ts` / `csvToWorkbookData`), now that the ceiling is Excel-parity.
 * Guards against a regression that would clamp/truncate a big import back to the
 * old ~1,024/8,192 caps. Uses a 100k-row CSV (cheap to generate, exercises the
 * worker import path + the raised ceiling end-to-end).
 */
test('opens a 100k-row file without truncating, within budget', async ({ page }) => {
  test.setTimeout(120_000);

  await page.goto('/');
  await waitForUniver(page);

  const ROWS = 100_000;

  // Build a 100k-row CSV fixture (A=r{i}, plus two more columns).
  const fs = await import('node:fs');
  const fixture = '/tmp/casual-sheets-large.csv';
  {
    const lines: string[] = new Array(ROWS);
    for (let i = 0; i < ROWS; i++) lines[i] = `r${i},x${i},y${i}`;
    fs.writeFileSync(fixture, lines.join('\n'));
  }

  const startRows = await page.evaluate(() => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return (window.__univerAPI!.getActiveWorkbook()!.getActiveSheet() as any).getMaxRows() as number;
  });
  expect(startRows).toBeLessThan(ROWS); // grid starts far smaller than the file

  // File → Open → inject the fixture.
  const t0 = Date.now();
  await page.getByTestId('menubar-file').click();
  const [chooser] = await Promise.all([
    page.waitForEvent('filechooser'),
    page.getByTestId('menu-item-open').click(),
  ]);
  await chooser.setFiles(fixture);

  // Wait until the deep cell (last row, well past the old caps) is mounted.
  await page.waitForFunction(
    (n) => {
      const api = window.__univerAPI;
      if (!api) return false;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()?.getActiveSheet();
      try {
        return ws?.getRange(n - 1, 0).getValue() === `r${n - 1}`;
      } catch {
        return false; // range still out of bounds → not loaded yet
      }
    },
    ROWS,
    { timeout: 90_000 },
  );
  const openMs = Date.now() - t0;
  console.log(`T1.4: opened ${ROWS}-row CSV in ${openMs} ms`);

  const after = await page.evaluate((n) => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    return {
      maxRows: ws.getMaxRows() as number,
      deep: ws.getRange(n - 1, 0).getValue(),
    };
  }, ROWS);

  // The full sheet loaded — not truncated to the old caps.
  expect(after.maxRows).toBeGreaterThanOrEqual(ROWS);
  expect(after.deep).toBe(`r${ROWS - 1}`);
});
