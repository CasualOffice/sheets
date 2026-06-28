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
 * Print cell-count cap — issue #50. The renderer builds one big HTML
 * table; past ~100k cells the tab OOMs. The fix surfaces a clear
 * error toast pointing at "Set Print Area" rather than crashing.
 *
 * We force the used range to ~200k cells by setting a value at a
 * far-out cell (the sheet's last-row / last-col detection picks it
 * up), then trigger Print and assert the toast.
 */
test('print on a too-large used range surfaces a clear toast rather than crashing', async ({
  page,
}) => {
  // Lower the limit before the app boots so we don't have to
  // synthesize a 100k-cell workbook in the test. The override hook
  // lives in `apps/web/src/shell/print.ts` for exactly this reason.
  await page.addInitScript(`window.__casualPrintCellLimit = 16;`);

  await page.goto('/');
  await waitForUniver(page);

  // Default sheet is 1024 × 26 — but the print path uses the
  // *used* range. Seed cells at A1 + somewhere far enough that
  // last-row × last-col > 16. 6×6 = 36 cells, comfortably above
  // the test cap and within the default sheet bounds.
  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange(0, 0).setValue({ v: 'tl' });
    ws.getRange(5, 5).setValue({ v: 'br' });
  });

  // Open File menu → Print, which opens the page-setup dialog.
  await page.getByTestId('menubar-file').click();
  await page.getByTestId('menu-item-print').click();

  // Confirm in the dialog. The button label is "Print".
  await page.getByRole('button', { name: /^Print$/i }).click();

  // Soft-guard fires → clear toast with the cell count + a pointer
  // to Set Print Area. Not a crash, not a silent failure.
  await expect(page.getByText(/Too many cells to print at once/i)).toBeVisible({
    timeout: 10_000,
  });
  await expect(page.getByText(/Set Print Area/i)).toBeVisible();
});
