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

import { expect, test, type Page } from '@playwright/test';
import { waitForUniver } from './_helpers';

/**
 * Macros — Phase 5, T5.1 (record) + T5.2 (run). Data → Macros records the
 * command-bus mutations your edits produce into a named macro, then replays
 * them. This drives the full loop through the real menu: record → edit → stop →
 * clear → run → the edits reappear.
 */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
type AnyApi = any;

const cell = (page: Page, a1: string) =>
  page.evaluate(
    (ref) =>
      (window.__univerAPI as AnyApi).getActiveWorkbook().getActiveSheet().getRange(ref).getValue(),
    a1,
  );
const setCell = (page: Page, a1: string, v: unknown) =>
  page.evaluate(
    ([ref, val]) =>
      (window.__univerAPI as AnyApi)
        .getActiveWorkbook()
        .getActiveSheet()
        .getRange(ref)
        .setValue({ v: val }),
    [a1, v] as const,
  );

async function macroItem(page: Page, testid: string) {
  await page.getByTestId('menubar-data').click();
  await page.getByTestId('menu-item-macros').hover();
  await page.getByTestId(testid).click();
}

test('record a macro, then replay it onto cleared cells', async ({ page }) => {
  test.setTimeout(60_000);
  await page.addInitScript(() => localStorage.removeItem('casual.macros'));
  await page.goto('/');
  await waitForUniver(page);

  // Record → edit A1/B1 → stop (saves "Macro 1").
  await macroItem(page, 'menu-item-macro-record');
  await setCell(page, 'A1', 5);
  await setCell(page, 'B1', 6);
  await page.waitForTimeout(200);
  await macroItem(page, 'menu-item-macro-record'); // now "Stop recording"

  // Clear the cells.
  await setCell(page, 'A1', '');
  await setCell(page, 'B1', '');
  await expect.poll(() => cell(page, 'A1')).toBeFalsy();

  // Run the saved macro → values come back.
  await macroItem(page, 'menu-item-macro-run-Macro-1');
  await expect.poll(() => cell(page, 'A1')).toBe(5);
  expect(await cell(page, 'B1')).toBe(6);
});

test('manage macros dialog runs and deletes a saved macro', async ({ page }) => {
  test.setTimeout(60_000);
  await page.addInitScript(() => localStorage.removeItem('casual.macros'));
  await page.goto('/');
  await waitForUniver(page);

  // Record → edit A1 → stop (saves "Macro 1").
  await macroItem(page, 'menu-item-macro-record');
  await setCell(page, 'A1', 7);
  await page.waitForTimeout(200);
  await macroItem(page, 'menu-item-macro-record'); // now "Stop recording"

  // Clear, then run via the Manage Macros dialog → value returns.
  await setCell(page, 'A1', '');
  await expect.poll(() => cell(page, 'A1')).toBeFalsy();
  await macroItem(page, 'menu-item-macro-manage');
  await expect(page.getByTestId('macros-dialog')).toBeVisible();
  await page.getByTestId('macros-dialog-run-Macro-1').click();
  await expect.poll(() => cell(page, 'A1')).toBe(7);

  // Reopen → delete the macro → list is empty, menu entry gone.
  await macroItem(page, 'menu-item-macro-manage');
  await page.getByTestId('macros-dialog-delete-Macro-1').click();
  await expect(page.getByTestId('macros-dialog-empty')).toBeVisible();
  await page.keyboard.press('Escape');
  await page.getByTestId('menubar-data').click();
  await page.getByTestId('menu-item-macros').hover();
  await expect(page.getByTestId('menu-item-macro-manage')).toHaveCount(0);
});

test('rename a macro from the manage dialog updates the run menu entry', async ({ page }) => {
  test.setTimeout(60_000);
  await page.addInitScript(() => localStorage.removeItem('casual.macros'));
  await page.goto('/');
  await waitForUniver(page);

  // Record "Macro 1".
  await macroItem(page, 'menu-item-macro-record');
  await setCell(page, 'A1', 4);
  await page.waitForTimeout(200);
  await macroItem(page, 'menu-item-macro-record'); // Stop

  // Open the dialog and rename Macro 1 → "Monthly close" (Enter commits).
  await macroItem(page, 'menu-item-macro-manage');
  const nameInput = page.getByTestId('macros-dialog-name-Macro-1');
  await nameInput.fill('Monthly close');
  await nameInput.press('Enter');
  // Row re-keys to the new name → run button reflects it; old slug is gone.
  await expect(page.getByTestId('macros-dialog-run-Monthly-close')).toBeVisible();
  await expect(page.getByTestId('macros-dialog-run-Macro-1')).toHaveCount(0);

  // Close → the Macros menu's quick-run entry uses the new name.
  await page.keyboard.press('Escape');
  await page.getByTestId('menubar-data').click();
  await page.getByTestId('menu-item-macros').hover();
  await expect(page.getByTestId('menu-item-macro-run-Monthly-close')).toBeVisible();
  await expect(page.getByTestId('menu-item-macro-run-Macro-1')).toHaveCount(0);
});

test('bind a macro to Ctrl+Shift+<letter> and trigger it from the keyboard', async ({ page }) => {
  test.setTimeout(60_000);
  await page.addInitScript(() => localStorage.removeItem('casual.macros'));
  await page.goto('/');
  await waitForUniver(page);

  // Record → edit A1 → stop (saves "Macro 1").
  await macroItem(page, 'menu-item-macro-record');
  await setCell(page, 'A1', 9);
  await page.waitForTimeout(200);
  await macroItem(page, 'menu-item-macro-record'); // now "Stop recording"

  // Bind it to Ctrl+Shift+M via the Manage Macros dialog.
  await macroItem(page, 'menu-item-macro-manage');
  await page.getByTestId('macros-dialog-shortcut-Macro-1').selectOption('M');
  await page.keyboard.press('Escape');

  // Clear A1, then fire the shortcut → the recorded value returns.
  await setCell(page, 'A1', '');
  await expect.poll(() => cell(page, 'A1')).toBeFalsy();
  await page.locator('body').click();
  await page.keyboard.press('Control+Shift+M');
  await expect.poll(() => cell(page, 'A1')).toBe(9);
});
