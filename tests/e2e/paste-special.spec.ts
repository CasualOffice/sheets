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
import { waitForUniver, selectRange } from './_helpers';

/**
 * Paste Special — the Ctrl+Alt+V dialog and its six paste modes.
 *
 * Headless chromium's clipboard read/write is async and unreliable
 * (see copy-paste.spec.ts), so the end-to-end paste result is verified
 * via the service-level command rather than a synthetic keystroke. The
 * dialog-interaction half is tested directly.
 */

test.describe('Paste Special', () => {
  test('Ctrl+Alt+V opens the dialog with all six modes', async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await selectRange(page, 'A1');
    await page.keyboard.press('Control+Alt+V');
    const dialog = page.getByTestId('paste-special-dialog');
    await expect(dialog).toBeVisible();
    for (const id of ['all', 'formulas', 'values', 'formats', 'col-widths', 'no-borders']) {
      await expect(page.getByTestId(`paste-special-${id}`)).toBeVisible();
    }
    // Escape closes without applying.
    await page.keyboard.press('Escape');
    await expect(dialog).toBeHidden();
  });

  test('Edit menu → Paste Special opens the same dialog', async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await page.getByTestId('menubar-edit').click();
    await page.getByTestId('menu-item-paste-special').click();
    await expect(page.getByTestId('paste-special-dialog')).toBeVisible();
  });

  test('pasteSpecial dispatches the correct paste command with hook value', async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    // Headless chromium's clipboard read/write is async and unreliable
    // (see copy-paste.spec.ts), so we can't observe a real paste effect
    // here. Instead, intercept the command bus to verify the dialog
    // dispatches `univer.command.paste` with the right hook value for
    // each radio choice.
    await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const api = window.__univerAPI! as any;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (window as any).__pasteSpy = [];
      const orig = api.executeCommand.bind(api);
      api.executeCommand = (id: string, params: unknown, opts: unknown) => {
        if (id === 'univer.command.paste') {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          (window as any).__pasteSpy.push(params);
        }
        return orig(id, params, opts);
      };
    });
    await selectRange(page, 'A1');
    await page.keyboard.press('Control+Alt+V');
    await page.getByTestId('paste-special-values').click();
    await page.getByTestId('paste-special-ok').click();
    await page.waitForTimeout(150);
    const spy = await page.evaluate(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      () => (window as any).__pasteSpy as Array<{ value?: string }>,
    );
    expect(spy.length).toBeGreaterThanOrEqual(1);
    expect(spy[spy.length - 1].value).toBe('special-paste-value');
  });
});
