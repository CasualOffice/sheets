/**
 * Copyright 2026 Casual Office
 *
 * Licensed under the Apache License, Version 2.0 (the "License").
 */

/**
 * Comments / Pivot / History panels on the SDK chrome rail. Extends the
 * foundation smoke: each rail button opens its panel without a runtime error
 * (the mount-crash guard — the real risk when a panel reaches a Univer service
 * or lazy plugin that isn't wired). Runs against the SDK chrome harness.
 */
import { expect, test } from '@playwright/test';

const PANELS = [
  { rail: 'pivot', body: 'cs-pivot-fields-panel' },
  { rail: 'comments', body: 'cs-comments-panel' },
  { rail: 'history', body: 'cs-history-panel' },
];

test.describe('SDK chrome light panels', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/sdk-harness?chrome=full');
    await page.waitForFunction(
      () => (window as unknown as { __sdkHarnessReady?: boolean }).__sdkHarnessReady === true,
      null,
      { timeout: 20_000 },
    );
    await expect(page.getByTestId('cs-menubar')).toBeVisible({ timeout: 20_000 });
  });

  for (const { rail, body } of PANELS) {
    test(`${rail} panel opens without a runtime error`, async ({ page }) => {
      const errors: string[] = [];
      page.on('pageerror', (e) => errors.push(e.message));

      await expect(page.getByTestId(`cs-panel-rail-${rail}`)).toBeVisible();
      await page.getByTestId(`cs-panel-rail-${rail}`).click();
      await expect(page.getByTestId(body)).toBeVisible({ timeout: 5_000 });

      expect(errors, `pageerror(s) opening ${rail}: ${errors.join(' | ')}`).toHaveLength(0);
      await page.screenshot({ path: `screenshots/panel-${rail}.png` });
    });
  }
});
