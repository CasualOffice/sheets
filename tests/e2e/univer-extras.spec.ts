import { expect, test } from '@playwright/test';
import { mainCanvas, waitForUniver } from './_helpers';

/**
 * Phase 0.5 — Univer OSS plugins we wired into the editor that ship in the
 * fork but weren't previously registered:
 *
 *   - sheets-crosshair-highlight: Excel-style active row/column highlight,
 *     toggled from the context-menu footer (off by default).
 *   - sheets-zen-editor: immersive full-screen cell editor, opened from the
 *     cell context menu.
 *
 * Both are eager plugins, so a clean editor boot already proves they register
 * without breaking Univer's DI / plugin order. These assertions additionally
 * prove each is reachable: zen via its context-menu entry, crosshair via its
 * toggle operation. A future Univer bump or plugin-list regression flags here.
 */

test.describe('Univer extras — crosshair highlight + zen editor', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
  });

  test('zen (full screen) editor is offered in the cell context menu', async ({ page }) => {
    const canvas = mainCanvas(page).first();
    await canvas.click({ position: { x: 80, y: 30 } });
    await canvas.click({ button: 'right', position: { x: 80, y: 30 } });
    const menu = page.locator('section.univer-popup').first();
    await expect(menu).toBeVisible({ timeout: 3_000 });
    await expect(menu).toContainText('Full Screen Editor');
  });

  test('crosshair highlight toggle operation is registered', async ({ page }) => {
    // Executes through the command bus; resolves false (not throws) only if the
    // operation is unregistered. A clean toggle there-and-back proves the plugin
    // is live without leaving the highlight on for other specs.
    const result = await page.evaluate(async () => {
      const api = window.__univerAPI!;
      const on = await api.executeCommand('sheet.operation.toggle-crosshair-highlight');
      const off = await api.executeCommand('sheet.operation.toggle-crosshair-highlight');
      return { on, off };
    });
    expect(result.on).toBe(true);
    expect(result.off).toBe(true);
  });
});
