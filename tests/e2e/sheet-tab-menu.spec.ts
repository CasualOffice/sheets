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
 * Right-click context menu on a sheet tab: Rename / Duplicate / Delete.
 *
 * Duplicate dispatches `sheet.command.copy-sheet` (provided by Univer's
 * sheets plugin); the result is a new tab whose name is the source name
 * with a Univer-default suffix. Rename and Delete reuse the existing
 * inline-rename input and deleteSheetById flow.
 */

test.describe('Sheet tab context menu', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
  });

  test('right-click opens the menu with Rename / Duplicate / Delete', async ({ page }) => {
    const firstTab = page.locator('[data-testid^="sheet-tab-"]').first();
    await firstTab.click({ button: 'right' });
    await expect(page.getByTestId('sheet-tab-context-menu')).toBeVisible();
    await expect(page.getByTestId('sheet-tab-menu-rename')).toBeVisible();
    await expect(page.getByTestId('sheet-tab-menu-duplicate')).toBeVisible();
    await expect(page.getByTestId('sheet-tab-menu-delete')).toBeVisible();
  });

  test('Escape closes the menu', async ({ page }) => {
    await page.locator('[data-testid^="sheet-tab-"]').first().click({ button: 'right' });
    await expect(page.getByTestId('sheet-tab-context-menu')).toBeVisible();
    await page.keyboard.press('Escape');
    await expect(page.getByTestId('sheet-tab-context-menu')).toBeHidden();
  });

  test('Duplicate creates a new sheet', async ({ page }) => {
    const initialCount = await page.evaluate(() => {
      const api = window.__univerAPI!;
      return api.getActiveWorkbook()!.getSheets().length;
    });

    await page.locator('[data-testid^="sheet-tab-"]').first().click({ button: 'right' });
    await page.getByTestId('sheet-tab-menu-duplicate').click();

    await page.waitForFunction(
      (n) => window.__univerAPI!.getActiveWorkbook()!.getSheets().length === n + 1,
      initialCount,
      { timeout: 3_000 },
    );
  });

  test('Delete is disabled when only one sheet remains', async ({ page }) => {
    await page.locator('[data-testid^="sheet-tab-"]').first().click({ button: 'right' });
    await expect(page.getByTestId('sheet-tab-menu-delete')).toBeDisabled();
  });

  test('Rename opens the inline editor on the right-clicked tab', async ({ page }) => {
    const firstTab = page.locator('[data-testid^="sheet-tab-"]').first();
    const tabId = await firstTab.getAttribute('data-testid');
    const sheetId = tabId!.replace('sheet-tab-', '');

    await firstTab.click({ button: 'right' });
    await page.getByTestId('sheet-tab-menu-rename').click();

    await expect(page.getByTestId(`sheet-tab-input-${sheetId}`)).toBeVisible();
  });
});
