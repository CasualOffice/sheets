import { expect, test } from '@playwright/test';
import { waitForUniver } from './_helpers';

/**
 * Freeze top row + Freeze first column should compose. Univer's built-in
 * set-first-row-frozen / set-first-column-frozen each zero the orthogonal
 * axis, so we route through the facade's additive setFrozenRows /
 * setFrozenColumns. Regression guard.
 */

type Freeze = { startRow: number; startColumn: number; ySplit: number; xSplit: number };

async function getFreeze(page: import('@playwright/test').Page): Promise<Freeze> {
  return page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const fws: any = api.getActiveWorkbook()!.getActiveSheet();
    return fws.getFreeze();
  });
}

test('Freeze top row then Freeze first column → both stay frozen', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);

  await page.getByTestId('menubar-view').click();
  await page.getByTestId('menu-item-freeze-row').click();

  let freeze = await getFreeze(page);
  expect(freeze.ySplit).toBe(1);
  expect(freeze.startRow).toBe(1);

  await page.getByTestId('menubar-view').click();
  await page.getByTestId('menu-item-freeze-col').click();

  freeze = await getFreeze(page);
  expect(freeze.ySplit).toBe(1);
  expect(freeze.startRow).toBe(1);
  expect(freeze.xSplit).toBe(1);
  expect(freeze.startColumn).toBe(1);
});

test('Freeze first column then Freeze top row → both stay frozen', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);

  await page.getByTestId('menubar-view').click();
  await page.getByTestId('menu-item-freeze-col').click();

  let freeze = await getFreeze(page);
  expect(freeze.xSplit).toBe(1);
  expect(freeze.startColumn).toBe(1);

  await page.getByTestId('menubar-view').click();
  await page.getByTestId('menu-item-freeze-row').click();

  freeze = await getFreeze(page);
  expect(freeze.ySplit).toBe(1);
  expect(freeze.startRow).toBe(1);
  expect(freeze.xSplit).toBe(1);
  expect(freeze.startColumn).toBe(1);
});
