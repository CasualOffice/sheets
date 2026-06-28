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
 * Auto-fit column width on a large sheet used to freeze the UI for 10s+.
 * `SheetSkeleton._calculateColWidth` measures every cell near the viewport
 * (±10 000 rows) and the fork used to build a full DocumentSkeleton per cell.
 * We added a fast path (plain value, no wrap, no rotation) that measures the
 * widest line with the LRU-cached FontCache — the same primitive the renderer
 * uses to size non-wrap content. This opens a 21k-row × 8-col file (the user's
 * real scenario) and locks in that auto-fitting all columns stays well under a
 * second; the bar is deliberately loose (CI is slower) but far below the old
 * 10s+ freeze.
 */
test('auto-fit column width on a 21k×8 sheet is fast', async ({ page }) => {
  test.setTimeout(120_000);

  await page.goto('/');
  await waitForUniver(page);

  const ROWS = 21_000;
  const COLS = 8;

  // Build the fixture: 8 columns, alternating wide text / numbers so auto-fit
  // has real work and the FontCache doesn't trivially hit on every cell.
  const fs = await import('node:fs');
  const fixture = '/tmp/casual-sheets-autofit.csv';
  {
    const lines: string[] = new Array(ROWS);
    for (let i = 0; i < ROWS; i++) {
      const cells: string[] = [];
      for (let c = 0; c < COLS; c++) {
        cells.push(c % 2 === 0 ? `Row ${i} value ${c}` : String(i * 100 + c));
      }
      lines[i] = cells.join(',');
    }
    fs.writeFileSync(fixture, lines.join('\n'));
  }

  // File → Open → inject the fixture (sets the true row extent, un-clamped).
  await page.getByTestId('menubar-file').click();
  const [chooser] = await Promise.all([
    page.waitForEvent('filechooser'),
    page.getByTestId('menu-item-open').click(),
  ]);
  await chooser.setFiles(fixture);

  // Wait until the deep cell is mounted (full sheet loaded).
  await page.waitForFunction(
    (n) => {
      const api = window.__univerAPI;
      if (!api) return false;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()?.getActiveSheet();
      try {
        return ws?.getRange(n - 1, 0).getValue() === `Row ${n - 1} value 0`;
      } catch {
        return false;
      }
    },
    ROWS,
    { timeout: 90_000 },
  );

  const result = await page.evaluate(async (cols) => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const api = (window as any).__univerAPI;
    const sheet = api.getActiveWorkbook().getActiveSheet();

    const t0 = performance.now();
    // setColumnAutoWidth(startCol, numCols) — the sheets-ui facade extension
    // the menu / Format-as-Table use; routes through the skeleton fast path.
    sheet.setColumnAutoWidth(0, cols);
    const ms = performance.now() - t0;

    return { ms, w0: sheet.getColumnWidth(0) as number };
  }, COLS);

  // eslint-disable-next-line no-console
  console.log(`auto-fit ${ROWS}×${COLS} took ${Math.round(result.ms)}ms, col0 width=${result.w0}`);
  expect(result.w0, 'auto-fit should widen the text column').toBeGreaterThan(20);
  expect(result.ms, 'auto-fit should not freeze the UI for seconds').toBeLessThan(3000);
});
