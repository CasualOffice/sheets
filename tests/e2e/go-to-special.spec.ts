import { expect, test } from '@playwright/test';
import { waitForUniver } from './_helpers';

/**
 * Go To Special (F5 / Ctrl+G / Edit → Go To Special…). Selects every cell on
 * the active sheet matching a criterion — constants, formulas, blanks, current
 * region, last cell. The selection maths is unit-tested in go-to-special.ts;
 * this drives the dialog end-to-end through the real app and asserts the
 * resulting selection lands on the right cells.
 */

async function seed(page: import('@playwright/test').Page) {
  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('A1').setValue({ v: 'Name' });
    ws.getRange('B1').setValue({ v: 'Score' });
    ws.getRange('A2').setValue({ v: 'Ada' });
    ws.getRange('B2').setValue({ v: 90 });
    ws.getRange('C2').setValue({ f: '=B2*2' });
    ws.getRange('A1').activate();
  });
}

async function activeCell(page: import('@playwright/test').Page) {
  return page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    const r = ws.getSelection().getActiveRange();
    return { row: r.getRow(), col: r.getColumn() };
  });
}

test('Edit → Go To Special selects formula cells', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);
  await seed(page);

  await page.getByTestId('menubar-edit').click();
  await page.getByTestId('menu-item-go-to-special').click();
  await expect(page.getByTestId('go-to-special-dialog')).toBeVisible();

  await page.getByTestId('go-to-special-formulas').check();
  await page.getByTestId('go-to-special-ok').click();
  await expect(page.getByTestId('go-to-special-dialog')).toBeHidden();

  // The only formula is C2 → the active cell jumps there.
  expect(await activeCell(page)).toEqual({ row: 1, col: 2 });
});

test('F5 opens the dialog and Constants moves the active cell off a formula', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);
  await seed(page);

  // Start the active cell on the formula, then F5 → Constants. (F5 is the Go To
  // Special shortcut; Ctrl+G focuses the Name Box for reference jumps.)
  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('C2').activate();
  });

  await page.keyboard.press('F5');
  await expect(page.getByTestId('go-to-special-dialog')).toBeVisible();
  await page.getByTestId('go-to-special-constants').check();
  await page.getByTestId('go-to-special-ok').click();
  await expect(page.getByTestId('go-to-special-dialog')).toBeHidden();

  // Constants are A1:B1 + A2:B2 — the active cell is the first constant (A1),
  // never the formula C2.
  expect(await activeCell(page)).toEqual({ row: 0, col: 0 });
});

test('Go To Special on blanks reports none when there are no gaps', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);
  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    // A solid 2x2 block — no blanks within the used range.
    ws.getRange('A1').setValue({ v: 1 });
    ws.getRange('B1').setValue({ v: 2 });
    ws.getRange('A2').setValue({ v: 3 });
    ws.getRange('B2').setValue({ v: 4 });
    ws.getRange('A1').activate();
  });

  await page.getByTestId('menubar-edit').click();
  await page.getByTestId('menu-item-go-to-special').click();
  await page.getByTestId('go-to-special-blanks').check();
  await page.getByTestId('go-to-special-ok').click();

  // No blanks → Excel's "No cells were found." and the dialog stays open.
  await expect(page.getByTestId('go-to-special-notice')).toContainText('No cells were found');
  await expect(page.getByTestId('go-to-special-dialog')).toBeVisible();
});
