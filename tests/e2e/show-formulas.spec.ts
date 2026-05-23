import { expect, test } from '@playwright/test';
import { waitForUniver } from './_helpers';

/**
 * Show Formulas (Ctrl+`) — non-destructive DOM overlay that paints
 * formula source text over every cell with a formula. Toggling off
 * hides the overlay; the underlying cell values are never mutated.
 */

test.describe('Show Formulas', () => {
  test('Ctrl+` toggles a per-formula overlay', async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('A1').setValue({ v: 10 });
      ws.getRange('A2').setValue({ v: 20 });
      ws.getRange('A3').setValue({ f: '=SUM(A1:A2)' });
    });
    // Layer hidden by default.
    await expect(page.getByTestId('show-formulas-layer')).toHaveCount(0);
    await page.keyboard.press('Control+`');
    await expect(page.getByTestId('show-formulas-layer')).toBeVisible({ timeout: 3_000 });
    // One overlay cell for the single formula.
    await expect(page.getByTestId('show-formulas-cell')).toHaveCount(1);
    await expect(page.getByTestId('show-formulas-cell').first()).toContainText('=SUM(A1:A2)');
    // Toggle off → overlay disappears.
    await page.keyboard.press('Control+`');
    await expect(page.getByTestId('show-formulas-layer')).toHaveCount(0);
  });

  test('View menu also toggles the overlay', async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('B1').setValue({ f: '=1+1' });
    });
    await page.getByTestId('menubar-view').click();
    await page.getByTestId('menu-item-show-formulas').click();
    await expect(page.getByTestId('show-formulas-cell').first()).toContainText('=1+1');
  });
});
