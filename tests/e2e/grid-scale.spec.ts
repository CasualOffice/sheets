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
import { readCell, selectRange, waitForUniver } from './_helpers';

/**
 * Grid-scale ceiling — Phase 1, T1.1a (docs/COMPETITIVE_ROADMAP.md).
 *
 * The interactive grid used to wall at MAX_ROWS = 8,192 (apps/web/src/snapshot.ts):
 * a user arrowing/scrolling down hit a hard stop. Raised to 65,536. The grid grows
 * on demand (useWorkbookGrowth appends a row chunk when the selection nears the
 * bottom edge), so this drives that hook with repeated Ctrl+ArrowDown and asserts
 * the sheet now grows past 8,192 — and that a cell beyond the old cap is editable.
 *
 * Lock-down: if MAX_ROWS regresses to <= 8,192, growth plateaus and this fails.
 */
test('grid grows past the old 8,192-row ceiling and stays editable', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);

  const getMaxRows = () =>
    page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      return ws.getMaxRows() as number;
    });

  // Start at A1 with the grid focused.
  await selectRange(page, 'A1');

  // Each Ctrl+ArrowDown on an empty column jumps to the grid's last row; the
  // growth hook appends a ~256-row chunk near that edge. ~28 presses cross 8,192
  // (1024 initial + n*256). Loop with early-exit; cap iterations as a backstop.
  let maxRows = await getMaxRows();
  for (let i = 0; i < 80 && maxRows <= 8192; i++) {
    await page.keyboard.press('Control+ArrowDown');
    await page.waitForTimeout(25);
    maxRows = await getMaxRows();
  }

  expect(maxRows, 'sheet should grow past the old 8,192-row cap').toBeGreaterThan(8192);

  // Jump to the new grid edge (now beyond row 8,192) so the cursor itself is
  // past the old cap — growth appends ahead of the cursor, so the loop can exit
  // with the selection still just under 8,192.
  await page.keyboard.press('Control+ArrowDown');
  const editRow = await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    const sel = ws.getActiveRange();
    return sel ? (sel.getRow() as number) : -1;
  });
  expect(editRow, 'cursor should reach a row past the old cap').toBeGreaterThan(8192);

  // That deep cell must accept and retain a value.
  await page.keyboard.type('past-8192');
  await page.keyboard.press('Enter');
  expect(await readCell(page, `A${editRow + 1}`)).toMatchObject({ v: 'past-8192' });
});
