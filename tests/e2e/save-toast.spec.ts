import { expect, test } from '@playwright/test';
import { waitForUniver } from './_helpers';

/**
 * Save confirmation toast. Each File → Save / Export path calls Univer's
 * IMessageService.show() with the resolved filename so the user gets in-app
 * feedback (the browser's download notification alone is too easy to miss).
 *
 * Univer renders the toast via Sonner under the `univer-message-toaster`
 * region — we assert by text rather than DOM id so the test is decoupled
 * from Sonner's internal markup.
 */
test('File → Save shows a "Saved as …" toast', async ({ page }) => {
  // Skip the actual browser download — clicking the anchor would trigger a
  // file save dialog in headless mode and time out the test. We only care
  // about the toast, which fires before/after the click regardless of
  // whether the browser persists the file.
  await page.addInitScript(() => {
    HTMLAnchorElement.prototype.click = function () {};
  });

  await page.goto('/');
  await waitForUniver(page);

  await page.getByTestId('menubar-file').click();
  await page.getByTestId('menu-item-save').click();

  await expect(page.getByText(/Saved as .+\.xlsx/i)).toBeVisible({
    timeout: 5_000,
  });
});
