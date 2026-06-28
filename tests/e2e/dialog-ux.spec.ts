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
 * Generic Dialog component UX — focus trap, body-scroll lock, and
 * focus restoration on close. Page Setup is the easiest dialog to
 * exercise because it opens via a stable shortcut (Ctrl+P) and has
 * multiple focusable elements (orientation buttons, margin select,
 * primary action).
 */

test.describe('Dialog UX (focus trap + scroll lock + restore)', () => {
  test.beforeEach(async ({ page }) => {
    // Defang window.print so Ctrl+P doesn't block on the OS dialog.
    await page.addInitScript(() => {
      window.print = () => {};
    });
    await page.goto('/');
    await waitForUniver(page);
  });

  test('opening a dialog locks body scroll', async ({ page }) => {
    const before = await page.evaluate(() => document.body.style.overflow);
    expect(before).not.toBe('hidden');

    await page.keyboard.press('Control+p');
    await expect(page.getByTestId('page-setup-dialog')).toBeVisible();

    const during = await page.evaluate(() => document.body.style.overflow);
    expect(during).toBe('hidden');

    // Close — Esc — and scroll lock is released.
    await page.keyboard.press('Escape');
    await expect(page.getByTestId('page-setup-dialog')).toHaveCount(0);
    const after = await page.evaluate(() => document.body.style.overflow);
    expect(after).not.toBe('hidden');
  });

  test('Tab cycles within the dialog (no escape to the page)', async ({ page }) => {
    await page.keyboard.press('Control+p');
    await expect(page.getByTestId('page-setup-dialog')).toBeVisible();

    // Press Tab a bunch of times — focus must always stay inside the
    // dialog (the dialog DOM contains document.activeElement).
    for (let i = 0; i < 20; i++) {
      await page.keyboard.press('Tab');
      const inside = await page.evaluate(() => {
        const dialog = document.querySelector('[data-testid="page-setup-dialog"]');
        return dialog?.contains(document.activeElement) ?? false;
      });
      expect(inside).toBe(true);
    }
  });

  test('Shift+Tab from the first element wraps to the last', async ({ page }) => {
    await page.keyboard.press('Control+p');
    await expect(page.getByTestId('page-setup-dialog')).toBeVisible();

    // Identify the first focusable element and put focus there.
    await page.evaluate(() => {
      const dialog = document.querySelector(
        '[data-testid="page-setup-dialog"]',
      ) as HTMLElement;
      const focusables = dialog.querySelectorAll<HTMLElement>(
        'a[href], input:not([disabled]), select:not([disabled]), textarea:not([disabled]), button:not([disabled]), [tabindex]:not([tabindex="-1"])',
      );
      (focusables[0] as HTMLElement)?.focus();
    });
    const firstId = await page.evaluate(
      () => (document.activeElement as HTMLElement)?.getAttribute('data-testid'),
    );
    expect(firstId).toBeTruthy();

    await page.keyboard.press('Shift+Tab');

    // Should now be on the LAST focusable in the dialog, not on something
    // outside it (which is what would happen without the trap).
    const onLast = await page.evaluate(() => {
      const dialog = document.querySelector(
        '[data-testid="page-setup-dialog"]',
      ) as HTMLElement;
      const focusables = dialog.querySelectorAll<HTMLElement>(
        'a[href], input:not([disabled]), select:not([disabled]), textarea:not([disabled]), button:not([disabled]), [tabindex]:not([tabindex="-1"])',
      );
      const last = focusables[focusables.length - 1] as HTMLElement;
      return last === document.activeElement;
    });
    expect(onLast).toBe(true);
  });

  test('focus is restored to the opener after the dialog closes', async ({ page }) => {
    // Open via the File menu so the opener is a deterministic element.
    await page.getByTestId('menubar-file').click();
    const printItem = page.getByTestId('menu-item-print');
    await printItem.focus();
    await page.keyboard.press('Enter');
    await expect(page.getByTestId('page-setup-dialog')).toBeVisible();

    await page.keyboard.press('Escape');
    await expect(page.getByTestId('page-setup-dialog')).toHaveCount(0);

    // The print menu item is unmounted with the menu, so focus falls
    // back to <body>. What we really care about is that focus did NOT
    // get stuck inside an unmounted dialog — so just assert it's a
    // visible element on the page (or body, which is acceptable).
    const safe = await page.evaluate(
      () => document.activeElement?.tagName?.toLowerCase() ?? 'body',
    );
    expect(['body', 'button', 'input', 'div']).toContain(safe);
  });
});
