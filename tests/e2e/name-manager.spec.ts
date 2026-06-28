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
 * Name Manager dialog — Excel's Ctrl+F3 for managing named ranges.
 * Verifies the create / list / delete loop using Univer's defined-name
 * facade as the source of truth.
 */

test.describe('Name Manager', () => {
  test('Ctrl+F3 opens the dialog with an empty state', async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await page.keyboard.press('Control+F3');
    const dialog = page.getByTestId('name-manager-dialog');
    await expect(dialog).toBeVisible();
    await expect(dialog).toContainText(/No named ranges yet|Refers to/);
    await page.keyboard.press('Escape');
    await expect(dialog).toBeHidden();
  });

  test('Data menu → Name Manager opens the same dialog', async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await page.getByTestId('menubar-data').click();
    await page.getByTestId('menu-item-name-manager').click();
    await expect(page.getByTestId('name-manager-dialog')).toBeVisible();
  });

  test('Create + delete a named range', async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await page.keyboard.press('Control+F3');
    await page.getByTestId('name-manager-new').click();
    await page.getByTestId('name-manager-name-input').fill('SalesQ3');
    await page.getByTestId('name-manager-ref-input').fill('Sheet1!$A$1:$A$10');
    await page.getByTestId('name-manager-save').click();
    // After save, the row should be in the table.
    const row = page.getByTestId('name-manager-row').first();
    await expect(row).toContainText('SalesQ3');
    await expect(row).toContainText('A$1');
    // Verify Univer sees it.
    const facadeSeesIt = await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const wb = api.getActiveWorkbook() as any;
      const list = wb?.getDefinedNames?.() ?? [];
      return list.some((d: { getName(): string }) => d.getName() === 'SalesQ3');
    });
    expect(facadeSeesIt).toBe(true);
    // Delete and verify the row disappears.
    await page.getByTestId('name-manager-delete').first().click();
    await expect(page.getByTestId('name-manager-row')).toHaveCount(0);
  });

  test('Create a worksheet-scoped name', async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);

    const sheetName = await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws = api.getActiveWorkbook()!.getActiveSheet() as any;
      return ws.getSheetName() as string;
    });

    await page.keyboard.press('Control+F3');
    await page.getByTestId('name-manager-new').click();
    await page.getByTestId('name-manager-name-input').fill('LocalRange');
    await page.getByTestId('name-manager-ref-input').fill(`${sheetName}!$A$1:$A$5`);
    // Scope it to the active sheet (not Workbook).
    await page.getByTestId('name-manager-scope-select').selectOption({ label: sheetName });
    await page.getByTestId('name-manager-save').click();

    // The row shows the sheet name in the Scope column.
    const scopeCell = page.getByTestId('name-manager-scope').first();
    await expect(scopeCell).toHaveText(sheetName);

    // Univer stored it as sheet-scoped (localSheetId = the sheet's id, not the
    // workbook sentinel).
    const scoped = await page.evaluate((name) => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const wb = api.getActiveWorkbook() as any;
      const sheetId = wb.getActiveSheet().getSheetId();
      const dn = wb.getDefinedName(name);
      return dn?.getLocalSheetId?.() === sheetId;
    }, 'LocalRange');
    expect(scoped).toBe(true);
  });

  test('Create rejects blank name', async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await page.keyboard.press('Control+F3');
    await page.getByTestId('name-manager-new').click();
    // Leave name blank, fill ref, click save → should error.
    await page.getByTestId('name-manager-ref-input').fill('Sheet1!$A$1');
    await page.getByTestId('name-manager-save').click();
    await expect(page.getByTestId('name-manager-error')).toBeVisible();
  });
});
