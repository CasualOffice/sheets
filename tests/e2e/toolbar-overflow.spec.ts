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
 * Toolbar overflow indicators — when the toolbar is wider than the
 * viewport, left/right chevron buttons appear at either edge so the
 * user can tell more buttons exist. Previously the row scrolled
 * silently with a hidden scrollbar.
 */

test.describe('Toolbar overflow chevrons', () => {
  test.beforeEach(async ({ page }) => {
    // Narrow viewport so the toolbar is guaranteed to overflow. 600 px
    // is comfortably below any plausible groups-row total width across
    // future ribbon adjustments.
    await page.setViewportSize({ width: 600, height: 800 });
    await page.goto('/');
    await waitForUniver(page);
  });

  test('right chevron appears when content overflows; clicking scrolls', async ({ page }) => {
    const right = page.getByTestId('toolbar-overflow-right');
    await expect(right).toBeVisible();

    // Clicking the chevron should move the scroller to the right.
    const before = await page.evaluate(
      () => (document.querySelector('.toolbar__inner') as HTMLElement).scrollLeft,
    );
    await right.click();
    // Smooth scroll is async — give it a tick before reading.
    await page.waitForTimeout(400);
    const after = await page.evaluate(
      () => (document.querySelector('.toolbar__inner') as HTMLElement).scrollLeft,
    );
    expect(after).toBeGreaterThan(before);
  });

  test('left chevron appears only after scrolling away from the start', async ({ page }) => {
    // Initially the scroller is at 0 — no left chevron yet.
    await expect(page.getByTestId('toolbar-overflow-left')).toHaveCount(0);

    // Programmatically scroll right.
    await page.evaluate(() => {
      const el = document.querySelector('.toolbar__inner') as HTMLElement;
      el.scrollLeft = 300;
    });
    await expect(page.getByTestId('toolbar-overflow-left')).toBeVisible();
  });

  test('right chevron disappears once scrolled to the end', async ({ page }) => {
    await page.evaluate(() => {
      const el = document.querySelector('.toolbar__inner') as HTMLElement;
      el.scrollLeft = el.scrollWidth;
    });
    await expect(page.getByTestId('toolbar-overflow-right')).toHaveCount(0);
  });
});
