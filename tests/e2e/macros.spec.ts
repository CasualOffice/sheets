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
