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
import { waitForUniver } from './_helpers';

/**
 * Page Setup dialog — opens before File → Print / Ctrl+P with orientation +
 * margin presets, then defers to the browser's native print dialog. Options
 * persist via localStorage so the next print starts from the same defaults.
 */

test.describe('Page Setup', () => {
  test.beforeEach(async ({ page }) => {
    // No-op the eventual window.print so the test never blocks on a real
    // print dialog. Has to be installed before the iframe attaches.
    await page.addInitScript(() => {
      window.print = () => {};
      const origAppend = document.body.appendChild.bind(document.body);
      document.body.appendChild = function <T extends Node>(node: T): T {
        const ret = origAppend(node);
        if (node instanceof HTMLIFrameElement) {
          node.addEventListener('load', () => {
            if (node.contentWindow) node.contentWindow.print = () => {};
          });
        }
        return ret;
      } as typeof document.body.appendChild;
    });

    await page.goto('/');
    await waitForUniver(page);
  });

  test('File → Print opens the Page Setup dialog with defaults', async ({ page }) => {
    await page.getByTestId('menubar-file').click();
    await page.getByTestId('menu-item-print').click();

    await expect(page.getByTestId('page-setup-dialog')).toBeVisible();
    // Defaults: portrait + normal margins (per DEFAULT_PRINT_OPTIONS).
    // The visual "selected" state lives on the label's --active modifier.
    await expect(page.getByTestId('page-setup-orientation-portrait')).toHaveClass(/page-setup__option--active/);
    await expect(page.getByTestId('page-setup-margins')).toHaveValue('normal');
  });

  test('Ctrl+P opens the same Page Setup dialog', async ({ page }) => {
    // Focus the grid host so the keystroke doesn't get swallowed.
    await page.locator('[data-testid="grid-host"]').click();
    await page.keyboard.press('Control+P');
    await expect(page.getByTestId('page-setup-dialog')).toBeVisible();
  });

  test('Picking Landscape bakes "landscape" into the printed page CSS', async ({ page }) => {
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('A1').setValue({ v: 'hello' });
    });

    await page.getByTestId('menubar-file').click();
    await page.getByTestId('menu-item-print').click();
    await page.getByTestId('page-setup-orientation-landscape').click();
    await page.getByTestId('page-setup-margins').selectOption('wide');
    await page.getByTestId('page-setup-print').click();

    // The iframe srcdoc carries the @page rule with the chosen orientation
    // and margin (wide = 25mm). If the dialog or print() ever stop threading
    // options through, this test catches it.
    const srcdoc = await page.waitForFunction(
      () => {
        const f = document.querySelector('iframe');
        return f?.srcdoc ?? null;
      },
      null,
      { timeout: 3000 },
    );
    const html = (await srcdoc.jsonValue()) as string;
    expect(html).toContain('@page { size: A4 landscape; margin: 25mm; }');
  });

  test('Options persist in localStorage and the next open reflects them', async ({ page }) => {
    await page.getByTestId('menubar-file').click();
    await page.getByTestId('menu-item-print').click();
    await page.getByTestId('page-setup-orientation-landscape').click();
    await page.getByTestId('page-setup-margins').selectOption('narrow');
    await page.getByTestId('page-setup-print').click();

    // Re-open — defaults should now be landscape + narrow.
    await page.getByTestId('menubar-file').click();
    await page.getByTestId('menu-item-print').click();
    await expect(page.getByTestId('page-setup-orientation-landscape')).toHaveClass(/page-setup__option--active/);
    await expect(page.getByTestId('page-setup-margins')).toHaveValue('narrow');
  });

  test('Cancel closes the dialog without printing', async ({ page }) => {
    await page.getByTestId('menubar-file').click();
    await page.getByTestId('menu-item-print').click();
    await page.getByTestId('page-setup-cancel').click();
    await expect(page.getByTestId('page-setup-dialog')).toBeHidden();
    // No print iframe should have been added.
    expect(await page.locator('iframe').count()).toBe(0);
  });
});
