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

import { expect, test, chromium, type Browser } from '@playwright/test';
import { waitForUniver } from './_helpers';

/**
 * Repro for the user-reported bug: peer A clicks "Add sheet", peer B
 * doesn't see the new tab until B removes one of the existing sheets.
 *
 * Asserts:
 *   - peer A's local sheet count grows by one
 *   - peer B's sheet count grows by one (this is the failing case)
 *   - peer B's tab strip shows the new tab
 */

const PROD_BASE = process.env.PROD_BASE ?? 'http://localhost:3000';

let browser: Browser | null = null;
test.describe.configure({ mode: 'serial' });

test.beforeAll(async () => {
  browser = await chromium.launch();
});
test.afterAll(async () => {
  await browser?.close();
});

function installEnv(name: string): string {
  return `
    (function () {
      try {
        localStorage.setItem('casual.collab.displayName', ${JSON.stringify(name)});
        localStorage.setItem('casual.collab.namePrompted', '1');
      } catch (_) {}
    })();
  `;
}

test('peer A adds sheet → peer B sees the new tab', async () => {
  const ownerCtx = await browser!.newContext();
  const owner = await ownerCtx.newPage();
  await owner.addInitScript({ content: installEnv('Alice') });
  await owner.goto(PROD_BASE);
  await waitForUniver(owner);
  const roomId = await owner.evaluate(async () => {
    const res = await fetch('/api/rooms', {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: '{}',
    });
    return ((await res.json()) as { roomId: string }).roomId;
  });
  await owner.goto(`${PROD_BASE}/r/${roomId}`);
  await waitForUniver(owner);
  await expect(owner.getByTestId('presence-avatars')).toBeVisible({ timeout: 10_000 });

  const joinerCtx = await browser!.newContext();
  const joiner = await joinerCtx.newPage();
  await joiner.addInitScript({ content: installEnv('Bob') });
  await joiner.goto(`${PROD_BASE}/r/${roomId}`);
  await waitForUniver(joiner);
  await expect(joiner.getByTestId('presence-avatars')).toBeVisible({ timeout: 10_000 });

  // Both peers start with 1 sheet.
  const startA = await owner.evaluate(() => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const api = (window as any).__univerAPI;
    return api.getActiveWorkbook().getSheets().length;
  });
  const startB = await joiner.evaluate(() => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const api = (window as any).__univerAPI;
    return api.getActiveWorkbook().getSheets().length;
  });
  expect(startA).toBe(1);
  expect(startB).toBe(1);

  // Owner adds a sheet via the actual addSheet button (mirrors the UI).
  await owner.getByTestId('sheet-tabs-add').click();
  await owner.waitForTimeout(500);

  const afterA = await owner.evaluate(() => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const api = (window as any).__univerAPI;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const sheets = api.getActiveWorkbook().getSheets() as any[];
    return { count: sheets.length, ids: sheets.map((s) => s.getSheetId()) };
  });
  expect(afterA.count, 'owner should see 2 sheets').toBe(2);

  // The peer-side propagation needs a moment (Yjs round-trip).
  await joiner.waitForTimeout(2000);

  const afterB = await joiner.evaluate(() => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const api = (window as any).__univerAPI;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const sheets = api.getActiveWorkbook().getSheets() as any[];
    return {
      count: sheets.length,
      ids: sheets.map((s) => s.getSheetId()),
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      logLen: (window as any).__hocuspocusProvider?.document?.getArray?.('ops')?.length,
    };
  });
  console.log('OWNER  sheets:', JSON.stringify(afterA));
  console.log('JOINER sheets:', JSON.stringify(afterB));

  expect(afterB.count, 'joiner should also see 2 sheets').toBe(2);
  expect(afterB.ids).toEqual(afterA.ids);

  // Tab strip in the DOM should reflect the new sheet. SheetTabs gives
  // each tab `data-testid="sheet-tab-<id>"`; per-id testid is more
  // specific than a prefix selector that would also catch
  // sheet-tab-input-*, sheet-tab-close-* etc.
  for (const id of afterB.ids) {
    await expect(joiner.locator(`[data-testid="sheet-tab-${id}"]`)).toBeVisible({
      timeout: 5_000,
    });
  }

  await ownerCtx.close();
  await joinerCtx.close();
});
