/**
 * Copyright 2026 Casual Office
 *
 * Licensed under the Apache License, Version 2.0 (the "License").
 */

/**
 * Charts panel + ECharts overlay in the SDK chrome. Opens the Charts panel from
 * the rail (which lazy-loads echarts — the mount-crash guard), then inserts a
 * chart over the selection and asserts the overlay renders on the grid.
 */
import { expect, test } from '@playwright/test';

test.describe('SDK chrome Charts', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/sdk-harness?chrome=full');
    await page.waitForFunction(
      () => (window as unknown as { __sdkHarnessReady?: boolean }).__sdkHarnessReady === true,
      null,
      { timeout: 20_000 },
    );
    await expect(page.getByTestId('cs-menubar')).toBeVisible({ timeout: 20_000 });
    await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const api = (window as any).__sdkHarnessAPI;
      const ws = api.univer.getActiveWorkbook().getActiveSheet();
      ws.getRange(0, 0).setValue('Q');
      ws.getRange(0, 1).setValue('Sales');
      ws.getRange(1, 0).setValue('Q1');
      ws.getRange(1, 1).setValue(120);
      ws.getRange(2, 0).setValue('Q2');
      ws.getRange(2, 1).setValue(90);
      ws.getRange(3, 0).setValue('Q3');
      ws.getRange(3, 1).setValue(150);
      ws.getRange('A1:B4').activate();
      await new Promise((r) => setTimeout(r, 150));
    });
  });

  test('Charts panel opens (echarts lazy-loads) with no runtime error', async ({ page }) => {
    const errors: string[] = [];
    page.on('pageerror', (e) => errors.push(e.message));
    await page.getByTestId('cs-panel-rail-charts').click();
    await expect(page.getByTestId('charts-panel')).toBeVisible({ timeout: 8_000 });
    expect(errors, `pageerror(s): ${errors.join(' | ')}`).toHaveLength(0);
    await page.screenshot({ path: 'screenshots/panel-charts.png' });
  });

  test('insert a chart → the overlay renders on the grid', async ({ page }) => {
    const errors: string[] = [];
    page.on('pageerror', (e) => errors.push(e.message));

    await page.getByTestId('cs-panel-rail-charts').click();
    await expect(page.getByTestId('charts-panel')).toBeVisible({ timeout: 8_000 });

    // Empty-state CTA (or the add button) opens the Insert Chart dialog.
    const cta = page.getByTestId('charts-panel-empty-cta');
    const add = page.getByTestId('charts-panel-add');
    if (await cta.isVisible().catch(() => false)) await cta.click();
    else await add.click();

    await expect(page.getByTestId('insert-chart-dialog')).toBeVisible({ timeout: 5_000 });
    await page.getByTestId('insert-chart-confirm').click();

    // The ECharts overlay paints on the grid.
    await expect(page.getByTestId('chart-overlay').first()).toBeVisible({ timeout: 8_000 });
    await expect(page.getByTestId('chart-overlay').first().locator('canvas').first()).toBeVisible({
      timeout: 8_000,
    });
    expect(errors, `pageerror(s): ${errors.join(' | ')}`).toHaveLength(0);
    await page.screenshot({ path: 'screenshots/chart-inserted.png' });
  });
});
