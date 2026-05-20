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
      const db: IDBDatabase = await new Promise((resolve, reject) => {
        const req = indexedDB.open('casual-sheets', 1);
        req.onupgradeneeded = () => {
          const d = req.result;
          if (!d.objectStoreNames.contains('autosave')) d.createObjectStore('autosave');
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

async function clearAutosaveDb(page: Page) {
  await page.evaluate(async () => {
    await new Promise<void>((resolve) => {
      const req = indexedDB.deleteDatabase('casual-sheets');
      req.onsuccess = req.onerror = req.onblocked = () => resolve();
    });
  });
}

test.describe('Autosave restore', () => {
  test.beforeEach(async ({ page }) => {
    // Land on the page once so we have an origin to talk to IDB on,
    // then wipe + reseed.
    await page.goto('/');
    await waitForUniver(page);
    await clearAutosaveDb(page);
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
});
