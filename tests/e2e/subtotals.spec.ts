import { expect, test } from '@playwright/test';
import { waitForUniver } from './_helpers';

/**
 * Subtotals (Data → Subtotal…). Inserts a SUBTOTAL row at each change in the
 * grouping column plus a grand total. Layout/formula maths is unit-tested in
 * subtotals.ts; this drives the dialog end-to-end and checks the inserted rows
 * + computed totals.
 */

test('inserts per-group subtotals and a grand total', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);

  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    const data: Array<[string, number]> = [
      ['Region', 0], // header (second col header set below)
      ['East', 10],
      ['East', 20],
      ['West', 30],
    ];
    data.forEach((row, r) => {
      ws.getRange(r, 0).setValue({ v: row[0] });
      ws.getRange(r, 1).setValue({ v: r === 0 ? 'Amount' : row[1] });
    });
    ws.getRange('A1:B4').activate();
  });

  await page.getByTestId('menubar-data').click();
  await page.getByTestId('menu-item-subtotal').click();
  await expect(page.getByTestId('subtotals-dialog')).toBeVisible();

  // Group by Region (col 0, default), Sum (default), subtotal Amount (col 1 → checked by default).
  await page.getByTestId('subtotals-ok').click();
  await expect(page.getByTestId('subtotals-dialog')).toBeHidden();

  // Wait for the formula worker to evaluate the SUBTOTAL cells — the labels
  // write synchronously but the SUBTOTAL values land asynchronously.
  await page.waitForFunction(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    // Grand Total (row 6, col B) resolves to 60 once recalc settles.
    return ws.getRange(6, 1).getValue() === 60;
  });

  const out = await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    const labels: unknown[] = [];
    const amounts: unknown[] = [];
    for (let r = 0; r < 7; r++) {
      labels.push(ws.getRange(r, 0).getCellData()?.v ?? null);
      amounts.push(ws.getRange(r, 1).getValue());
    }
    return { labels, amounts };
  });

  expect(out.labels).toEqual([
    'Region',
    'East',
    'East',
    'East Total',
    'West',
    'West Total',
    'Grand Total',
  ]);
  // East Total = 30, West Total = 30, Grand Total = 60.
  expect(out.amounts[3]).toBe(30);
  expect(out.amounts[5]).toBe(30);
  expect(out.amounts[6]).toBe(60);
});

test('warns when the selection is too small', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);
  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('A1').setValue({ v: 'lonely' });
    ws.getRange('A1').activate();
  });

  await page.getByTestId('menubar-data').click();
  await page.getByTestId('menu-item-subtotal').click();
  await expect(page.getByTestId('subtotals-notice')).toBeVisible();
  await expect(page.getByTestId('subtotals-ok')).toBeDisabled();
});
