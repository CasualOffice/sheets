/**
 * Activity log — UX_AUDIT.md §4.1 / Phase 4 #14.
 *
 * Drives the end-to-end shape with synthetic errors fired through the
 * window-event bridge (same channel ToastContext uses). Avoids
 * coupling the spec to any specific real failure surface — the
 * bridge is the contract.
 */
import { test, expect } from '@playwright/test';
import { waitForUniver } from './_helpers';

test.describe('Activity log', () => {
  test('idle: no pill rendered when there are no entries', async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await expect(page.getByTestId('activity-pill')).toHaveCount(0);
  });

  test('error event surfaces a badge; popover lists it; dismiss removes it', async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);

    // Fire two synthetic error events through the same bridge the
    // ToastContext uses on every toast.error call.
    await page.evaluate(() => {
      window.dispatchEvent(
        new CustomEvent('cd:activity-error', {
          detail: { message: 'Save failed: network down' },
        }),
      );
      window.dispatchEvent(
        new CustomEvent('cd:activity-error', {
          detail: { message: 'Open failed: bad xlsx' },
        }),
      );
    });

    // Pill appears with badge "2".
    await expect(page.getByTestId('activity-pill')).toBeVisible();
    const badge = page.getByTestId('activity-pill-badge');
    await expect(badge).toHaveText('2');

    // Open the popover; the badge clears (markAllRead).
    await page.getByTestId('activity-pill-trigger').click();
    await expect(page.getByTestId('activity-pill-popover')).toBeVisible();
    await expect(page.getByText('Save failed: network down')).toBeVisible();
    await expect(page.getByText('Open failed: bad xlsx')).toBeVisible();
    // Badge gone once read.
    await expect(page.getByTestId('activity-pill-badge')).toHaveCount(0);

    // Dismiss one entry — the other survives.
    const firstDismiss = page.locator('[data-testid$="-dismiss"]').first();
    await firstDismiss.click();
    // 2 → 1 entry; the popover header reflects it.
    await expect(page.getByTestId('activity-pill-popover')).toContainText('1 entry');
  });

  test('Clear all empties the log and removes the pill', async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);

    await page.evaluate(() => {
      window.dispatchEvent(
        new CustomEvent('cd:activity-error', {
          detail: { message: 'Boom' },
        }),
      );
    });

    await page.getByTestId('activity-pill-trigger').click();
    await page.getByTestId('activity-clear-all').click();

    // Pill self-hides once the log is empty.
    await expect(page.getByTestId('activity-pill')).toHaveCount(0);
  });
});
