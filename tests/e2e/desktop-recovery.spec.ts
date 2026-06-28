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
 * Desktop crash-recovery — the native-app analogue of autosave. The Rust
 * shell owns a per-file recovery sidecar; the editor writes a debounced
 * snapshot to it, reads it on open, and offers to restore.
 *
 * The shell bridge (`window.__deskApp__`) only exists inside the Tauri app,
 * so here we inject a fake one before the page loads. The bootstrap is a
 * complete no-op without `?desk=1`, so the injected bridge survives. The
 * fake persists its "cleared" state in localStorage so the reload assertion
 * (discard must not re-prompt) exercises the real read-after-clear path.
 */

const RECORD = {
  name: 'Recovered workbook',
  sourceFormat: null as string | null,
  data: {
    id: 'wb-recovery-test',
    rev: 1,
    name: 'Recovered workbook',
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
};

async function injectRecoveryBridge(page: Page, record: typeof RECORD) {
  await page.addInitScript((rec) => {
    // localStorage survives reloads within the same context, so a Discard
    // (which sets the flag) keeps the sidecar "gone" on the next load.
    const CLEARED = '__deskapp_recovery_cleared__';
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (window as any).__deskApp__ = {
      isDesktop: true,
      filePath: '/tmp/recovered.xlsx',
      async readRecovery() {
        if (localStorage.getItem(CLEARED) === '1') return null;
        return new TextEncoder().encode(JSON.stringify(rec)).buffer;
      },
      async writeRecovery() {
        /* no-op for the test */
      },
      async clearRecovery() {
        localStorage.setItem(CLEARED, '1');
      },
    };
  }, record);
}

test.describe('Desktop crash-recovery', () => {
  test('restore banner appears and applies the recovered snapshot', async ({ page }) => {
    await injectRecoveryBridge(page, RECORD);
    await page.goto('/');
    await waitForUniver(page);

    const banner = page.getByTestId('desktop-recovery-banner');
    await expect(banner).toBeVisible();
    await expect(banner).toContainText('Recovered workbook');

    await page.getByTestId('desktop-recovery-restore').click();
    await expect(banner).toHaveCount(0);

    const a1 = await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      return ws.getRange(0, 0).getValue();
    });
    expect(a1).toBe('restored');
  });

  test('discard clears the sidecar and the banner does not return on reload', async ({ page }) => {
    await injectRecoveryBridge(page, RECORD);
    await page.goto('/');
    await waitForUniver(page);

    await page.getByTestId('desktop-recovery-discard').click();
    await expect(page.getByTestId('desktop-recovery-banner')).toHaveCount(0);

    // Reload — clearRecovery set the cleared flag, so readRecovery returns
    // null and the banner stays gone.
    await page.reload();
    await waitForUniver(page);
    await expect(page.getByTestId('desktop-recovery-banner')).toHaveCount(0);
  });
});
