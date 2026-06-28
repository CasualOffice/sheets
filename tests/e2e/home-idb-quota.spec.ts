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
 * Mode 1 (browser-only) IDB quota banner — shown when
 * `navigator.storage.estimate()` reports the visitor is over the soft
 * threshold (50 MB). Tracked as a Phase A follow-up under #49 and
 * shipped alongside Phase C close-out (2026-06-06).
 *
 * We stub `navigator.storage.estimate` via `addInitScript` so the
 * test doesn't depend on the visitor's real disk usage.
 */

test.describe('idb quota banner', () => {
  test('hidden when usage is under the threshold', async ({ page }) => {
    await page.addInitScript(`
      Object.defineProperty(navigator, 'storage', {
        value: { estimate: async () => ({ usage: 1024 * 1024, quota: 100 * 1024 * 1024 }) },
        configurable: true,
      });
    `);
    await page.goto('/');
    await waitForUniver(page, { keepHome: true });
    await expect(page.getByTestId('idb-quota-banner')).toHaveCount(0);
  });

  test('renders + can be dismissed once usage crosses 50 MB', async ({ page }) => {
    await page.addInitScript(`
      Object.defineProperty(navigator, 'storage', {
        value: {
          estimate: async () => ({ usage: 60 * 1024 * 1024, quota: 500 * 1024 * 1024 }),
        },
        configurable: true,
      });
    `);
    await page.goto('/');
    await waitForUniver(page, { keepHome: true });

    const banner = page.getByTestId('idb-quota-banner');
    await expect(banner).toBeVisible({ timeout: 10_000 });
    await expect(banner).toContainText('~60 MB');

    await page.getByTestId('idb-quota-dismiss').click();
    await expect(banner).toHaveCount(0);

    // Dismiss flag persists in sessionStorage.
    const flag = await page.evaluate(() =>
      sessionStorage.getItem('casual-sheets:idb-quota-banner-dismissed'),
    );
    expect(flag).toBe('1');
  });

  test('absent when navigator.storage.estimate is missing', async ({ page }) => {
    await page.addInitScript(`
      Object.defineProperty(navigator, 'storage', {
        value: undefined,
        configurable: true,
      });
    `);
    await page.goto('/');
    await waitForUniver(page, { keepHome: true });
    await expect(page.getByTestId('idb-quota-banner')).toHaveCount(0);
  });
});
