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
