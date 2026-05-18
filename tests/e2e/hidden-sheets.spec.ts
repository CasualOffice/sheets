import { expect, test, type Page } from '@playwright/test';
import { waitForUniver } from './_helpers';

/**
 * Hidden-sheet UX — Excel keeps hidden sheets in the workbook but
 * removes them from the tab strip. Without an inline affordance the
 * user has no way to bring them back; the tabs panel didn't surface
 * hidden state at all. Now the tab strip shows a small badge button
 * when at least one sheet is hidden, and clicking it opens a menu
 * to show any of them.
 */

async function addSheet(page: Page, name: string) {
  await page.evaluate((name) => {
    const api = window.__univerAPI!;
    const wb = api.getActiveWorkbook()!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (wb as any).insertSheet();
    const sheets = wb.getSheets();
    const last = sheets[sheets.length - 1];
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (last as any).setName(name);
  }, name);
}

async function hideSheetByName(page: Page, name: string) {
  await page.evaluate((name) => {
    const api = window.__univerAPI!;
    const wb = api.getActiveWorkbook()!;
    const target = wb.getSheets().find((s) => s.getSheetName() === name);
    if (!target) throw new Error('not found: ' + name);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (target as any).hideSheet();
  }, name);
}

test.describe('Hidden sheets', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await addSheet(page, 'Secret');
    await addSheet(page, 'Public');
  });

  test('a hidden sheet disappears from the tab strip', async ({ page }) => {
    const list = page.getByTestId('sheet-tabs-list');
    await expect(list).toContainText('Secret');
    await hideSheetByName(page, 'Secret');
    await expect(list).not.toContainText('Secret');
    // Other tabs unaffected.
    await expect(list).toContainText('Public');
    await expect(list).toContainText('Sheet1');
  });

  test('the hidden-sheets button appears with a badge count', async ({ page }) => {
    await expect(page.getByTestId('sheet-tabs-hidden')).toHaveCount(0);
    await hideSheetByName(page, 'Secret');
    const btn = page.getByTestId('sheet-tabs-hidden');
    await expect(btn).toBeVisible();
    await expect(btn).toContainText('1');
  });

  test('clicking a hidden sheet in the menu restores it to the tab strip', async ({ page }) => {
    await hideSheetByName(page, 'Secret');
    await page.getByTestId('sheet-tabs-hidden').click();
    const menu = page.getByTestId('hidden-sheets-menu');
    await expect(menu).toBeVisible();
    await expect(menu).toContainText('Secret');
    await menu.getByRole('menuitem', { name: /Secret/i }).click();
    // Restored.
    await expect(page.getByTestId('sheet-tabs-list')).toContainText('Secret');
    // Badge button disappears once nothing is hidden.
    await expect(page.getByTestId('sheet-tabs-hidden')).toHaveCount(0);
  });

  test('right-click → Hide sheet hides the active tab', async ({ page }) => {
    // Right-click on the Public tab.
    const publicTab = page.locator('.sheet-tab', { hasText: 'Public' });
    await publicTab.click({ button: 'right' });
    const menu = page.getByTestId('sheet-tab-context-menu');
    await expect(menu).toBeVisible();
    await menu.getByTestId('sheet-tab-menu-hide').click();
    await expect(page.getByTestId('sheet-tabs-list')).not.toContainText('Public');
    await expect(page.getByTestId('sheet-tabs-hidden')).toBeVisible();
  });

  test('Hide is disabled when only one visible sheet remains', async ({ page }) => {
    // Hide Public and Sheet1 so only Secret is visible.
    await hideSheetByName(page, 'Public');
    await hideSheetByName(page, 'Sheet1');
    const remaining = page.locator('.sheet-tab', { hasText: 'Secret' });
    await remaining.click({ button: 'right' });
    await expect(page.getByTestId('sheet-tab-menu-hide')).toBeDisabled();
  });
});
