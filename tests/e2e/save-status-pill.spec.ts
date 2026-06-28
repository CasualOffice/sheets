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

/**
 * Save-status pill — UX_AUDIT.md §4.3 / Phase 4 #16.
 *
 * Drives the lifecycle through a real save:
 *  - Idle by default (no pill rendered).
 *  - Ctrl+S → 'Saved' pill appears (browser save path completes
 *    synchronously enough that 'Saving…' is not reliably observable
 *    in headless without a delay; we assert the terminal state).
 *  - Typing after a Save knocks the pill back to idle (markDirty).
 */
import { test, expect } from '@playwright/test';
import { waitForUniver } from './_helpers';

test.describe('SaveStatusPill', () => {
  test('Ctrl+S surfaces "Saved" then a fresh edit clears it', async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    // Idle by default — pill is intentionally absent when there's
    // nothing to say (avoids cluttering the title bar on first paint).
    await expect(page.getByTestId('save-status-pill')).toHaveCount(0);

    // Need a real content edit so the empty-draft guard doesn't short-circuit Save.
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('A1').setValue({ v: 'hello' });
    });
    // Suppress the browser-source save's download blob — saveAs in
    // the test environment downloads the workbook, which Playwright
    // intercepts as a navigation. Cancelling the click on `<a>` is
    // overkill; suppress the download via a route-level no-op.
    await page.context().setDefaultTimeout(8000);
    const [download] = await Promise.all([
      page.waitForEvent('download').catch(() => null),
      page.keyboard.press('Control+s'),
    ]);
    if (download) await download.cancel();

    // The pill flips to "Saved …" — relative time can vary, so just
    // assert the prefix.
    const pill = page.getByTestId('save-status-pill');
    await expect(pill).toHaveAttribute('data-state', 'saved');
    await expect(pill).toContainText(/Saved/);

    // Typing again drops the pill back to idle (markDirty effect).
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('A2').setValue({ v: 'world' });
    });
    await expect(pill).toHaveCount(0);
  });
});
