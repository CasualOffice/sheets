import { expect, test } from '@playwright/test';
import { waitForUniver } from './_helpers';

/**
 * Busy pill — title-bar indicator the app surfaces while a known-slow
 * action is mid-flight (Format-as-Table on big sheets, Sort, Fill).
 * It's the UX patch for "looks frozen" until we modify Univer for
 * actual perf.
 *
 * We test the pill plumbing directly via a dev-only window hook
 * rather than via a real slow command — small e2e fixtures complete
 * the work too fast for Playwright's expect polling to catch the
 * brief visibility window deterministically. The real Format-as-
 * Table wiring (Toolbar.tsx + TablesPanel.tsx) calls the same
 * setBusy/runBusy that the spec drives here.
 */
test.describe('Busy indicator pill', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
  });

  test('shows the label when setBusy is called and hides on clear', async ({ page }) => {
    const pill = page.getByTestId('busy-pill');
    await expect(pill).toHaveCount(0);

    await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (globalThis as any).__setBusy?.('Creating table…');
    });

    await expect(pill).toBeVisible({ timeout: 2_000 });
    await expect(pill).toHaveText(/Creating table/);

    await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (globalThis as any).__setBusy?.(null);
    });
    await expect(pill).toBeHidden({ timeout: 2_000 });
  });

  test('the pill is pointer-events: none so it does not block clicks', async ({ page }) => {
    await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (globalThis as any).__setBusy?.('Working…');
    });
    const pill = page.getByTestId('busy-pill');
    await expect(pill).toBeVisible();
    const pointerEvents = await pill.evaluate((el) => getComputedStyle(el).pointerEvents);
    expect(pointerEvents).toBe('none');
  });
});
