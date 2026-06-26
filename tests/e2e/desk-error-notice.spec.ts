import { test, expect } from '@playwright/test';

/**
 * In desktop mode an unexpected runtime error should surface as a compact,
 * dismissible notice — not the old full-width red monospace banner across the
 * grid (which read like a crash/dev artifact). Detail goes to the console.
 *
 * The bridge installs its error listeners at module load, so we can drive a
 * synthetic error without waiting for the whole grid to boot.
 */
test('desktop runtime errors show a compact dismissible notice, not a red banner', async ({
  page,
}) => {
  await page.goto('/?desk=1');

  await page.evaluate(() => {
    window.dispatchEvent(
      new ErrorEvent('error', { message: 'boom', filename: 'x.js', lineno: 1, colno: 1 }),
    );
  });

  const notice = page.locator('#__deskapp_err__');
  await expect(notice).toBeVisible();
  await expect(notice).toHaveAttribute('role', 'alert');
  await expect(notice).toContainText('Something went wrong');
  // The raw error text is NOT dumped into the UI (it goes to the console).
  await expect(notice).not.toContainText('boom');

  // Dismissible.
  await notice.getByRole('button', { name: 'Dismiss' }).click();
  await expect(notice).toHaveCount(0);
});
