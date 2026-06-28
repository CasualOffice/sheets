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
 * Autosave + restore — catches the "I closed the tab without saving"
 * failure mode without standing up a backend.
 *
 *   - Driver writes a snapshot to IndexedDB on idle / on pagehide.
 *   - On next open, a restore banner offers to reapply the snapshot.
 *   - Explicit Save clears the slot so it doesn't re-prompt.
 *   - Skipped inside /r/<id> co-edit rooms (covered by the server).
 *
 * The driver debounces 5s; the test seeds IDB directly so we don't have
 * to wait — the same path the driver writes through (key 'current' in
 * store 'autosave' of DB 'casual-sheets').
 */

async function seedAutosave(page: Page, name: string) {
  await page.evaluate(
    async ({ wbName }) => {
      // Open at the same version the app uses (autosave + version-history +
      // recent-files all share this DB). Seeding at an older version
      // would fail with VersionError on the next app boot and silently
      // kill autosave reads.
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
        const tx = db.transaction('autosave', 'readwrite');
        tx.objectStore('autosave').put(
          {
            name: wbName,
            sourceFormat: null,
            data: {
              id: 'wb-autosave-test',
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
            savedAt: Date.now() - 1000,
          },
          'current',
        );
        tx.oncomplete = () => resolve();
        tx.onerror = () => reject(tx.error);
      });
    },
    { wbName: name },
  );
}

test.describe('Autosave restore', () => {
  test.beforeEach(async ({ page }) => {
    // Land on the page once so we have an origin to talk to IDB on.
    // Each Playwright test runs in its own BrowserContext with a fresh
    // IDB, so no explicit wipe is needed. The pre-Phase-A version of
    // this beforeEach did a `deleteDatabase` here; that races with the
    // page's autosave / recent-files / pinned-folder hooks holding
    // the v4 schema open and hangs when its delete blocks.
    await page.goto('/');
    await waitForUniver(page);
  });

  test('restore banner appears on next load and applies the saved snapshot', async ({ page }) => {
    await seedAutosave(page, 'Recovered draft');
    await page.goto('/');
    await waitForUniver(page);

    const banner = page.getByTestId('autosave-banner');
    await expect(banner).toBeVisible();
    await expect(banner).toContainText('Recovered draft');
    await page.getByTestId('autosave-restore').click();
    await expect(banner).toHaveCount(0);

    const a1 = await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      return ws.getRange(0, 0).getValue();
    });
    expect(a1).toBe('restored');
  });

  test('discard drops the slot and the banner does not return on reload', async ({ page }) => {
    await seedAutosave(page, 'Should be discarded');
    await page.goto('/');
    await waitForUniver(page);
    await page.getByTestId('autosave-discard').click();
    await expect(page.getByTestId('autosave-banner')).toHaveCount(0);

    // Reload — no autosave record exists, so no banner.
    await page.reload();
    await waitForUniver(page);
    await expect(page.getByTestId('autosave-banner')).toHaveCount(0);
  });

  // Exercises the LIVE driver write end-to-end (not the seeded path): a real
  // edit must still reach IndexedDB now that the snapshot is deferred to an
  // idle slot. Guards against the idle scheduling silently dropping the save.
  test('a live edit is written to the autosave slot (idle-scheduled)', async ({ page }) => {
    test.setTimeout(30_000);
    // A real interaction is required before the driver will mark dirty.
    await page.mouse.click(200, 200);
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange(0, 0).setValue({ v: 'live-autosave' });
    });

    // The driver debounces 5s, then writes on idle — poll the slot until the
    // edit lands in IndexedDB.
    await expect
      .poll(
        () =>
          page.evaluate(
            () =>
              new Promise<unknown>((resolve) => {
                const req = indexedDB.open('casual-sheets', 4);
                req.onsuccess = () => {
                  const db = req.result;
                  const get = db
                    .transaction('autosave', 'readonly')
                    .objectStore('autosave')
                    .get('current');
                  get.onsuccess = () => {
                    const rec = get.result as
                      | {
                          data?: {
                            sheetOrder?: string[];
                            sheets?: Record<
                              string,
                              { cellData?: Record<string, Record<string, { v?: unknown }>> }
                            >;
                          };
                        }
                      | undefined;
                    const sid = rec?.data?.sheetOrder?.[0];
                    db.close();
                    resolve(sid ? rec?.data?.sheets?.[sid]?.cellData?.['0']?.['0']?.v : null);
                  };
                  get.onerror = () => {
                    db.close();
                    resolve(null);
                  };
                };
                req.onerror = () => resolve(null);
              }),
          ),
        { timeout: 20_000, intervals: [500] },
      )
      .toBe('live-autosave');
  });
});
