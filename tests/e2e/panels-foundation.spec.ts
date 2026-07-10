/**
 * Copyright 2026 Casual Office
 *
 * Licensed under the Apache License, Version 2.0 (the "License").
 */

/**
 * Panel foundation (rail + host + registry) in the SDK chrome. The standalone
 * app's panel rail never existed in the SDK that embedders (dochub) use; this
 * locks in that `chrome="full"` now renders the rail and that the Tables panel
 * opens from it.
 */
import { expect, test } from '@playwright/test';

test.describe('SDK chrome panel rail', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/sdk-harness?chrome=full');
    await page.waitForFunction(
      () => (window as unknown as { __sdkHarnessReady?: boolean }).__sdkHarnessReady === true,
      null,
      { timeout: 20_000 },
    );
    await expect(page.getByTestId('cs-menubar')).toBeVisible({ timeout: 20_000 });
  });

  test('rail is visible with the Tables button', async ({ page }) => {
    await expect(page.getByTestId('cs-panel-rail')).toBeVisible();
    await expect(page.getByTestId('cs-panel-rail-tables')).toBeVisible();
  });

  test('clicking Tables opens the panel; clicking again closes it', async ({ page }) => {
    const errors: string[] = [];
    page.on('pageerror', (e) => errors.push(e.message));

    await page.getByTestId('cs-panel-rail-tables').click();
    await expect(page.getByTestId('cs-tables-panel')).toBeVisible({ timeout: 5_000 });
    await expect(page.getByTestId('cs-tables-panel-empty')).toBeVisible();
    await page.screenshot({ path: 'screenshots/panel-tables.png' });

    await page.getByTestId('cs-panel-rail-tables').click();
    await expect(page.getByTestId('cs-tables-panel')).toHaveCount(0);

    expect(errors, `pageerror(s): ${errors.join(' | ')}`).toHaveLength(0);
  });
});
