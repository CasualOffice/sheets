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

test.describe('Sheet tabs', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
  });

  test('Renders the Add button on the left and a tab for the initial sheet', async ({ page }) => {
    await expect(page.getByTestId('sheet-tabs')).toBeVisible();
    await expect(page.getByTestId('sheet-tabs-add')).toBeVisible();

    const tabs = page.getByTestId('sheet-tabs').locator('.sheet-tab');
    await expect(tabs).toHaveCount(1);
    await expect(tabs.first()).toContainText('Sheet1');
  });

  test('Plus button adds a new sheet', async ({ page }) => {
    await page.getByTestId('sheet-tabs-add').click();
    const tabs = page.getByTestId('sheet-tabs').locator('.sheet-tab');
    await expect(tabs).toHaveCount(2);
  });

  test('Clicking a tab switches the active sheet', async ({ page }) => {
    await page.getByTestId('sheet-tabs-add').click();
    const tabs = page.getByTestId('sheet-tabs').locator('.sheet-tab');

    await tabs.nth(1).click();
    await expect(tabs.nth(1)).toHaveAttribute('aria-selected', 'true');
    await expect(tabs.nth(0)).toHaveAttribute('aria-selected', 'false');
  });

  test('Double-click on a tab enables rename, Enter commits', async ({ page }) => {
    const tab = page.getByTestId('sheet-tabs').locator('.sheet-tab').first();
    await tab.dblclick();

    const input = page.locator('.sheet-tab__input').first();
    await expect(input).toBeVisible();
    await input.fill('Forecast');
    await input.press('Enter');

    await expect(tab).toContainText('Forecast');
  });

  test('× button on a tab deletes it', async ({ page }) => {
    await page.getByTestId('sheet-tabs-add').click();
    const tabs = page.getByTestId('sheet-tabs').locator('.sheet-tab');
    await expect(tabs).toHaveCount(2);

    const secondTabId = await tabs.nth(1).getAttribute('data-testid');
    expect(secondTabId).not.toBeNull();
    const sheetId = secondTabId!.replace('sheet-tab-', '');

    // The × is opacity:0 unless hovered/active — Playwright clicks work
    // regardless of visual opacity as long as the element is in the DOM.
    await page.getByTestId(`sheet-tab-close-${sheetId}`).click();

    await expect(page.getByTestId('sheet-tabs').locator('.sheet-tab')).toHaveCount(1);
  });

  test('Last remaining sheet × is disabled', async ({ page }) => {
    const tab = page.getByTestId('sheet-tabs').locator('.sheet-tab').first();
    const sheetId = (await tab.getAttribute('data-testid'))!.replace('sheet-tab-', '');
    await expect(page.getByTestId(`sheet-tab-close-${sheetId}`)).toBeDisabled();
  });

  test('moveSheetTo reorders the sheet via Facade (drag uses the same path)', async ({
    page,
  }) => {
    // The user-facing path is HTML5 drag-and-drop, which Playwright's
    // headless DnD support is unreliable for. We test the underlying
    // operation here; the React onDrop handler calls the same path.
    await page.getByTestId('sheet-tabs-add').click();
    await page.getByTestId('sheet-tabs-add').click();

    const tabs = page.getByTestId('sheet-tabs').locator('.sheet-tab');
    await expect(tabs).toHaveCount(3);

    // Move the last sheet to position 0 directly via Facade.
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      const wb = api.getActiveWorkbook()!;
      const sheets = wb.getSheets();
      const last = sheets[sheets.length - 1];
      wb.moveSheet(last, 0);
    });

    const order = await tabs.evaluateAll((els) => els.map((el) => el.textContent ?? ''));
    // The originally-last sheet (Sheet3) is now first.
    expect(order[0]).toContain('Sheet3');
  });
});
