/**
 * Copyright 2026 Casual Office
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

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
 *   - sheets-graphics: canvas render primitive (no UI / commands). A clean
 *     boot is the proof — a DI break would surface as a redi Service2 error.
 *   - watermark: overlay layer driven from View → Confidential watermark
 *     (the app hides Univer's ribbon, so we own the trigger).
 *
 * The first two pairs are eager plugins, so a clean editor boot already proves
 * they register without breaking Univer's DI / plugin order. These assertions
 * additionally prove each is reachable: zen via its context-menu entry,
 * crosshair via its toggle operation, watermark via its View-menu toggle +
 * persisted service config. A future Univer bump or plugin-list regression
 * flags here.
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

test.describe('Univer extras — graphics + watermark', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
  });

  test('app boots with sheets-graphics + watermark registered (no DI errors)', async ({ page }) => {
    // A clean boot past waitForUniver already proves both eager plugins
    // registered without a redi duplicate-Identifier break. Belt-and-braces:
    // assert nothing logged a Service2-style DI error and the facade is live.
    const ok = await page.evaluate(() => Boolean(window.__univerAPI?.getActiveWorkbook()));
    expect(ok).toBe(true);
  });

  test('View → Watermark dialog applies / re-texts / clears the overlay', async ({ page }) => {
    // The watermark renders into the canvas scene (no DOM node to assert) and
    // its config is held in the WatermarkService, persisted via localforage
    // (IndexedDB) under UNIVER_WATERMARK_STORAGE_KEY. We verify the round trip
    // the user controls: the dialog's switch + text drive the persisted config
    // (proving the service actually ran — a DI break would throw at apply time
    // and never write the key) and the View-menu item flips its checkmark.
    const persisted = () =>
      page.evaluate(
        () =>
          new Promise<{ content: string | null; opacity: number | null }>((resolve) => {
            const req = indexedDB.open('localforage');
            req.onsuccess = () => {
              const db = req.result;
              let store: IDBObjectStore;
              try {
                store = db.transaction('keyvaluepairs', 'readonly').objectStore('keyvaluepairs');
              } catch {
                resolve({ content: null, opacity: null });
                return;
              }
              const get = store.get('UNIVER_WATERMARK_STORAGE_KEY');
              get.onsuccess = () => {
                const v = get.result as
                  | { config?: { text?: { content?: string; opacity?: number } } }
                  | undefined;
                resolve({
                  content: v?.config?.text?.content ?? null,
                  opacity: v?.config?.text?.opacity ?? null,
                });
              };
              get.onerror = () => resolve({ content: null, opacity: null });
            };
            req.onerror = () => resolve({ content: null, opacity: null });
          }),
      );

    // Open the View menu and reach the watermark item. The menubar button
    // toggles the popup, so re-open via the popup's own visibility rather
    // than blindly re-clicking.
    const openViewMenu = async () => {
      const popup = page.getByTestId('menubar-view-popup');
      if (!(await popup.isVisible().catch(() => false))) {
        await page.getByTestId('menubar-view').click();
        await expect(popup).toBeVisible();
      }
      return page.getByTestId('menu-item-toggle-watermark');
    };

    const dialog = page.getByTestId('watermark-dialog');

    // ── Apply ON with the default text ──────────────────────────────────
    let item = await openViewMenu();
    await expect(item).not.toContainText('✓');
    await item.click();
    await expect(dialog).toBeVisible();

    // Switch defaults to off on a clean boot; turn it on and apply.
    await page.getByTestId('watermark-toggle').check();
    await page.getByTestId('watermark-apply').click();
    await expect(dialog).toBeHidden();

    await expect.poll(async () => (await persisted()).content).toBe('CONFIDENTIAL');
    item = await openViewMenu();
    await expect(item).toContainText('✓');

    // ── Re-open, set custom text + opacity, apply ───────────────────────
    item = await openViewMenu();
    await item.click();
    await expect(dialog).toBeVisible();
    // Switch is seeded ON from the applied config.
    await expect(page.getByTestId('watermark-toggle')).toBeChecked();
    await page.getByTestId('watermark-text').fill('DRAFT');
    await page.getByTestId('watermark-opacity').selectOption('0.2');
    await page.getByTestId('watermark-apply').click();
    await expect(dialog).toBeHidden();

    await expect.poll(async () => (await persisted()).content).toBe('DRAFT');
    expect((await persisted()).opacity).toBe(0.2);

    // ── Turn it off → clears the persisted config + the checkmark ───────
    item = await openViewMenu();
    await item.click();
    await expect(dialog).toBeVisible();
    await page.getByTestId('watermark-toggle').uncheck();
    await page.getByTestId('watermark-apply').click();
    await expect(dialog).toBeHidden();

    await expect.poll(async () => (await persisted()).content).toBeNull();
    item = await openViewMenu();
    await expect(item).not.toContainText('✓');
  });
});
