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
 * Selection stats at scale — Phase 1, T1.5 (docs/COMPETITIVE_ROADMAP.md).
 *
 * Raising the grid ceiling to Excel parity made a full-column / Ctrl+A selection
 * span up to 1,048,576 × 16,384 cells — which alone exceeds the 100k selection-
 * stats cap, so the status-bar stats would vanish (a regression: Excel shows
 * full-column sums). `useActiveCellState` now clamps the selection to the used
 * range before materializing values, so stats stay correct AND fast (the cap
 * still guards a genuinely huge dense selection — no freeze).
 */
test('full-column selection on a 1M-row sheet still shows stats, fast', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);

  // Seed A1:A100 = 1..100 (sum 5050), then grow the grid so a full-column
  // selection spans ~1,000,000 cells — far past the 100k stats cap.
  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange(0, 0, 100, 1).setValues(Array.from({ length: 100 }, (_, i) => [i + 1]));
    ws.setRowCount(1_000_000);
  });

  await selectRange(page, 'A1');
  const t0 = Date.now();
  await page.keyboard.press('Control+Space'); // select the entire column A (~1M cells)

  // Stats must still appear (clamped to the used range). Without the clamp the
  // 1M-cell selection exceeds the cap and the Sum row is hidden.
  const sum = page.getByTestId('stat-sum');
  await expect(sum).toBeVisible({ timeout: 10_000 });
  const txt = (await sum.textContent()) ?? '';
  expect(Number(txt.replace(/[^0-9.-]/g, ''))).toBe(5050);

  // And it resolved quickly — no multi-second freeze materializing millions of cells.
  expect(Date.now() - t0).toBeLessThan(8_000);
});
