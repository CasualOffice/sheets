import { expect, test } from '@playwright/test';
import { mainCanvas, waitForUniver } from './_helpers';

/**
 * Charts P0 — ECharts plumbing. Locks in:
 *
 *   - Insert > Chart command inserts a ChartModel into the store.
 *   - ChartLayer renders an overlay for it anchored to the active
 *     sheet, positioned via the cell-rect math we share with
 *     PresenceLayer.
 *   - The overlay contains the ECharts canvas (renderer initialised).
 *   - Hidden when the chart's sheet isn't active.
 *
 * Persistence + drawing-model integration come in later phases.
 */

test.describe('Charts P0 — ECharts overlay', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);

    // Plant header + 3 rows of data so build-option has something
    // real to render: A1:C4.
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('A1').setValue({ v: 'Region' });
      ws.getRange('B1').setValue({ v: 'Q1' });
      ws.getRange('C1').setValue({ v: 'Q2' });
      ws.getRange('A2').setValue({ v: 'North' });
      ws.getRange('B2').setValue({ v: 100 });
      ws.getRange('C2').setValue({ v: 120 });
      ws.getRange('A3').setValue({ v: 'South' });
      ws.getRange('B3').setValue({ v: 80 });
      ws.getRange('C3').setValue({ v: 95 });
      ws.getRange('A4').setValue({ v: 'East' });
      ws.getRange('B4').setValue({ v: 110 });
      ws.getRange('C4').setValue({ v: 130 });
      ws.getRange('A1:C4').activate();
    });
  });

  test('Insert > Chart adds an overlay anchored to the active sheet', async ({ page }) => {
    // Click the canvas first so menu focus chain matches a real user.
    await mainCanvas(page).first().click({ position: { x: 100, y: 100 } });
    // Re-select after the canvas click.
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('A1:C4').activate();
    });

    await page.getByTestId('menubar-insert').click();
    await page.getByTestId('menu-item-insert-chart').click();

    // ChartLayer mounts an overlay per chart; ECharts seeds a canvas
    // inside. Both should exist within a couple of frames.
    const overlay = page.getByTestId('chart-overlay');
    await expect(overlay).toBeVisible({ timeout: 5_000 });
    await expect(overlay.locator('canvas')).toHaveCount(1, { timeout: 5_000 });

    // The overlay has a non-zero box from the cell-rect math.
    const box = await overlay.boundingBox();
    expect(box).not.toBeNull();
    expect(box!.width).toBeGreaterThan(50);
    expect(box!.height).toBeGreaterThan(50);
  });

  test('chart overlay hides when switching to a different sheet', async ({ page }) => {
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      const wb = api.getActiveWorkbook()!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = wb.getActiveSheet();
      ws.getRange('A1:C4').activate();
    });
    await page.getByTestId('menubar-insert').click();
    await page.getByTestId('menu-item-insert-chart').click();
    await expect(page.getByTestId('chart-overlay')).toBeVisible({ timeout: 5_000 });

    // Insert a second sheet + switch to it. Chart should disappear.
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      const wb = api.getActiveWorkbook()!;
      wb.insertSheet();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const sheets = wb.getSheets() as any[];
      wb.setActiveSheet(sheets[sheets.length - 1]);
    });
    await expect(page.getByTestId('chart-overlay')).toHaveCount(0, { timeout: 5_000 });
  });
});
