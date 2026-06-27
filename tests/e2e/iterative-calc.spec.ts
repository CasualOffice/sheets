import { expect, test } from '@playwright/test';
import { waitForUniver } from './_helpers';

/**
 * Iterative calculation toggle (Data → Iterative calculation). The engine
 * supports circular-reference convergence (cycle count > 1) but there was no
 * UI. With it on, a convergent self-reference like `=A1/2+5` settles to 10;
 * with it off (default cycle count 1) it stays a circular-reference error.
 */

test('enabling iterative calculation lets a circular formula converge', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);

  // Turn the toggle on first.
  await page.getByTestId('menubar-data').click();
  await page.getByTestId('menu-item-iterative-calc').click();

  // A convergent self-referential formula: x = x/2 + 5 → 10.
  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('A1').setValue({ f: '=A1/2+5' });
  });

  // It should converge to ~10 once the engine iterates.
  await page.waitForFunction(
    () => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      const v = ws.getRange('A1').getValue();
      return typeof v === 'number' && Math.abs(v - 10) < 0.5;
    },
    null,
    { timeout: 10_000 },
  );

  // The menu item now shows the enabled (✓) state.
  await page.getByTestId('menubar-data').click();
  await expect(page.getByTestId('menu-item-iterative-calc')).toContainText('✓');
});
