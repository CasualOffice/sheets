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
import { waitForUniver, selectRange } from './_helpers';

/**
 * Solo (single-user) version-history panel. Mirrors the co-edit
 * history panel but feeds off a local `ICommandService` subscription
 * so editing without a room still produces a changes record. The
 * collab variant is covered by `coedit-regression.spec.ts` already;
 * this spec asserts the solo path renders entries and supports revert.
 */

test.describe('Solo version history', () => {
  test('local edits show up in the History panel', async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);

    // Open the panel via the View menu (the only entry point besides
    // direct test wiring — keeps the test on the real user path).
    await page.getByTestId('menubar-view').click();
    await page.getByTestId('menu-item-history-panel').click();
    // The outer panel is the new VersionHistoryPanel (Versions tab by
    // default). The per-mutation log lives behind the Activity tab —
    // switch to it so the `history-panel` testid (the inner solo
    // history component) becomes visible.
    await expect(page.getByTestId('version-history-panel')).toBeVisible();
    await page.getByTestId('version-history-tab-activity').click();
    const panel = page.getByTestId('history-panel');
    await expect(panel).toBeVisible();

    // Solo session starts empty.
    await expect(panel.getByTestId('history-count')).toHaveText('0');

    await selectRange(page, 'A1');
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('A1').setValue({ v: 'hello' });
    });

    // The set-range-values mutation rides the command bus on the same
    // microtask; React batches the state update on the next frame.
    await expect(panel.getByTestId('history-count')).not.toHaveText('0');
    await expect(panel.getByTestId('history-row').first()).toContainText(/Edited/);
  });

  test('Revert undoes a local cell edit', async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await page.getByTestId('menubar-view').click();
    await page.getByTestId('menu-item-history-panel').click();
    // Switch to the Activity tab where the per-mutation rows live.
    await expect(page.getByTestId('version-history-panel')).toBeVisible();
    await page.getByTestId('version-history-tab-activity').click();
    await selectRange(page, 'B2');
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('B2').setValue({ v: 'before-revert' });
    });
    const firstRow = page.getByTestId('history-row').first();
    await expect(firstRow).toBeVisible();
    await firstRow.getByTestId('history-revert').click();
    await page.waitForTimeout(200);
    const v = await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      return ws.getRange('B2').getValue();
    });
    expect(v == null || v === '').toBeTruthy();
  });
});
