import { expect, test, type Page } from '@playwright/test';
import { waitForUniver } from './_helpers';

/**
 * Trace Precedents / Dependents (Data → Formula auditing). Draws tracer arrows
 * from a formula's precedent cells to it (and from a cell to its dependents).
 * Reference parsing is unit-tested in trace-model.ts; this drives the real menu
 * + overlay and asserts the right number of arrows render.
 */

async function activate(page: Page, a1: string) {
  await page.evaluate((ref) => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange(ref).activate();
  }, a1);
}

test.beforeEach(async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);
  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('A1').setValue({ v: 10 });
    ws.getRange('A2').setValue({ v: 20 });
    ws.getRange('C1').setValue({ f: '=A1+A2' });
  });
});

test('Trace Precedents draws an arrow from each precedent; Remove Arrows clears', async ({
  page,
}) => {
  await activate(page, 'C1');
  await page.getByTestId('menubar-data').click();
  await page.getByTestId('menu-item-trace-precedents').click();

  // Two precedents (A1, A2) → two arrows.
  await expect(page.getByTestId('trace-layer')).toBeVisible();
  await expect(page.getByTestId('trace-arrow')).toHaveCount(2);

  // Remove Arrows clears the overlay.
  await page.getByTestId('menubar-data').click();
  await page.getByTestId('menu-item-trace-clear').click();
  await expect(page.getByTestId('trace-arrow')).toHaveCount(0);
});

test('Trace Dependents draws an arrow to each dependent', async ({ page }) => {
  // A1 is referenced by C1 (=A1+A2) → one dependent.
  await activate(page, 'A1');
  await page.getByTestId('menubar-data').click();
  await page.getByTestId('menu-item-trace-dependents').click();
  await expect(page.getByTestId('trace-arrow')).toHaveCount(1);
});

test('Trace Precedents on a non-formula cell draws nothing', async ({ page }) => {
  await activate(page, 'A1'); // a constant
  await page.getByTestId('menubar-data').click();
  await page.getByTestId('menu-item-trace-precedents').click();
  await expect(page.getByTestId('trace-arrow')).toHaveCount(0);
});
