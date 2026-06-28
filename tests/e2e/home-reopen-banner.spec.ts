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

import { expect, test, type Page } from '@playwright/test';
import { waitForUniver } from './_helpers';

/**
 * Reopen banner — Phase A of the storage-modes work (issue #49).
 *
 * Cleanly-closed counterpart to the autosave banner: if a recent file
 * was opened within the last 7 days, the home screen offers a one-tap
 * "Pick up where you left off — <name>" prompt above the template
 * gallery.
 *
 * Tested by seeding the `recent-files` IDB store directly so we don't
 * have to round-trip through an actual open.
 */

async function seedRecent(page: Page, name: string, ageMs = 1000): Promise<void> {
  await page.evaluate(
    async ({ wbName, age }) => {
      const db: IDBDatabase = await new Promise((resolve, reject) => {
        const req = indexedDB.open('casual-sheets', 4);
        req.onupgradeneeded = () => {
          const d = req.result;
          if (!d.objectStoreNames.contains('autosave')) d.createObjectStore('autosave');
          if (!d.objectStoreNames.contains('versions')) {
            const os = d.createObjectStore('versions', { keyPath: 'id', autoIncrement: true });
            os.createIndex('savedAt', 'savedAt', { unique: false });
            os.createIndex('kind', 'kind', { unique: false });
          }
          if (!d.objectStoreNames.contains('recent-files')) {
            const os = d.createObjectStore('recent-files', { keyPath: 'id', autoIncrement: true });
            os.createIndex('openedAt', 'openedAt', { unique: false });
            os.createIndex('name', 'name', { unique: false });
          }
          if (!d.objectStoreNames.contains('pinned-folder')) {
            d.createObjectStore('pinned-folder');
          }
        };
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error);
      });
      await new Promise<void>((resolve, reject) => {
        const tx = db.transaction('recent-files', 'readwrite');
        tx.objectStore('recent-files').add({
          name: wbName,
          sourceFormat: 'xlsx',
          data: {
            id: 'wb-recent-test',
            rev: 1,
            name: wbName,
            appVersion: '0.22.1',
            locale: 1,
            styles: {},
            sheetOrder: ['s1'],
            sheets: {
              s1: {
                id: 's1',
                name: 'Sheet1',
                cellData: { '0': { '0': { v: 'restored' } } },
                rowCount: 100,
                columnCount: 26,
              },
            },
          },
          size: 256,
          openedAt: Date.now() - age,
        });
        tx.oncomplete = () => resolve();
        tx.onerror = () => reject(tx.error);
      });
      // Clear any dismiss flag from a previous run.
      try {
        sessionStorage.removeItem('casual-sheets:reopen-banner-dismissed');
        sessionStorage.removeItem('casual-sheets:reopen-banner-dismissed:id');
      } catch {
        /* sessionStorage blocked — fine */
      }
    },
    { wbName: name, age: ageMs },
  );
}

test.describe('home reopen banner', () => {
  test.beforeEach(async ({ page }) => {
    // Playwright gives each test a fresh BrowserContext with isolated
    // IDB — no wipe needed. We only need to land on the origin once so
    // subsequent `page.evaluate` calls have a same-origin context to
    // run in.
    await page.goto('/');
    await waitForUniver(page);
  });

  test('renders for a recently-opened file and clicking Open restores it', async ({ page }) => {
    await seedRecent(page, 'Quarterly Review');
    await page.goto('/');
    // Keep home up so the banner has a place to render.
    await waitForUniver(page, { keepHome: true });

    const banner = page.getByTestId('home-reopen-banner');
    await expect(banner).toBeVisible();
    await expect(banner).toContainText('Quarterly Review');

    await page.getByTestId('home-reopen-open').click();

    // The workbook swap dismisses the home and lifts the file name into
    // the title bar.
    await expect(page.getByTestId('home-screen')).toHaveCount(0, { timeout: 10_000 });
    await expect(page.getByRole('button', { name: 'Quarterly Review' })).toBeVisible();
  });

  test('dismiss hides the banner for the session', async ({ page }) => {
    await seedRecent(page, 'Daily standup notes');
    await page.goto('/');
    await waitForUniver(page, { keepHome: true });

    const banner = page.getByTestId('home-reopen-banner');
    await expect(banner).toBeVisible();

    await page.getByTestId('home-reopen-dismiss').click();
    await expect(banner).toHaveCount(0);

    // Re-render the home (no full reload — sessionStorage persists) by
    // toggling its visibility through the close-X path is overkill; we
    // just confirm the dismiss flag stuck.
    const flag = await page.evaluate(() =>
      sessionStorage.getItem('casual-sheets:reopen-banner-dismissed'),
    );
    expect(flag).toBe('1');
  });

  test('does not render for an entry older than 7 days', async ({ page }) => {
    await seedRecent(page, 'Old archive', 8 * 24 * 60 * 60 * 1000);
    await page.goto('/');
    await waitForUniver(page, { keepHome: true });

    // Recents strip itself still renders (stale gate is 60 days), but
    // the banner cut-off is 7.
    await expect(page.getByTestId('home-recent-open').first()).toBeVisible();
    await expect(page.getByTestId('home-reopen-banner')).toHaveCount(0);
  });
});
