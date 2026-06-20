import { expect, test } from '@playwright/test';

/**
 * Find & Replace chrome dialog (packages/sdk/src/chrome/FindReplace.tsx),
 * exercised via the `/sdk-harness` route. Kept in its own spec so it doesn't
 * collide with sdk-harness.spec.ts when batches land in parallel.
 */

test.describe('SDK chrome — Find & Replace', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/sdk-harness?chrome=minimal');
    await page.waitForFunction(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      () => (window as any).__sdkHarnessReady === true,
      null,
      { timeout: 30_000 },
    );
    await page.waitForFunction(
      () => Array.from(document.querySelectorAll('canvas')).some((c) => c.clientWidth > 0),
      null,
      { timeout: 30_000 },
    );
    // Seed A1=apple, A2=banana, A3=apricot.
    await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const api = (window as any).__sdkHarnessAPI;
      const ws = api.univer.getActiveWorkbook().getActiveSheet();
      ws.getRange(0, 0).setValue('apple');
      ws.getRange(1, 0).setValue('banana');
      ws.getRange(2, 0).setValue('apricot');
      ws.getRange('A1').activate();
      await new Promise((r) => setTimeout(r, 150));
    });
  });

  test('Ctrl+F opens the dialog and counts matches', async ({ page }) => {
    await page.keyboard.press('Control+f');
    await expect(page.getByTestId('cs-find-replace')).toBeVisible();
    await page.getByTestId('cs-find-input').fill('ap');
    // "apple" + "apricot" match; "banana" doesn't.
    await expect(page.getByTestId('cs-find-count')).toHaveText('1/2');
    // Escape closes.
    await page.getByTestId('cs-find-input').press('Escape');
    await expect(page.getByTestId('cs-find-replace')).toHaveCount(0);
  });

  test('match case narrows results', async ({ page }) => {
    await page.keyboard.press('Control+f');
    await page.getByTestId('cs-find-input').fill('AP');
    // Case-insensitive by default → 2 matches.
    await expect(page.getByTestId('cs-find-count')).toHaveText('1/2');
    await page.getByTestId('cs-find-match-case').check();
    // Case-sensitive "AP" matches nothing (cells are lowercase).
    await expect(page.getByTestId('cs-find-count')).toHaveText('0/0');
  });

  test('Replace All rewrites every match', async ({ page }) => {
    await page.keyboard.press('Control+h');
    await expect(page.getByTestId('cs-replace-input')).toBeVisible();
    await page.getByTestId('cs-find-input').fill('ap');
    await page.getByTestId('cs-replace-input').fill('X');
    await page.getByTestId('cs-replace-all').click();
    const vals = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const api = (window as any).__sdkHarnessAPI;
      const ws = api.univer.getActiveWorkbook().getActiveSheet();
      for (let i = 0; i < 20; i++) {
        const a1 = ws.getRange(0, 0).getValue();
        if (a1 === 'Xple') break;
        await new Promise((r) => setTimeout(r, 100));
      }
      return {
        a1: ws.getRange(0, 0).getValue(),
        a2: ws.getRange(1, 0).getValue(),
        a3: ws.getRange(2, 0).getValue(),
      };
    });
    expect(vals.a1).toBe('Xple');
    expect(vals.a2).toBe('banana');
    expect(vals.a3).toBe('Xricot');
  });
});
