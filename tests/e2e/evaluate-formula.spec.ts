import { expect, test, type Page } from '@playwright/test';
import { waitForUniver } from './_helpers';

/**
 * Evaluate Formula (Data → Evaluate Formula…). Steps through the active cell's
 * formula one sub-expression at a time, substituting each computed value, until
 * a single result remains — and leaves the cell's formula untouched (the engine
 * calculates off-cell). Stepping is unit-tested in formula-evaluate.ts; this
 * drives the real dialog + `executeFormulas`.
 */

async function cellFormula(page: Page, a1: string) {
  return page.evaluate((ref) => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    return ws.getRange(ref).getFormula();
  }, a1);
}

test('steps through a nested formula and leaves the cell intact', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);

  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('A1').setValue({ v: 3 });
    ws.getRange('A2').setValue({ v: 4 });
    ws.getRange('D7').setValue({ v: 9 });
    ws.getRange('C1').setValue({ f: '=SUM(A1:A2)+SQRT(D7)' });
    ws.getRange('C1').activate();
  });

  await page.getByTestId('menubar-data').click();
  await page.getByTestId('menu-item-evaluate-formula').click();
  const expr = page.getByTestId('evaluate-formula-expr');
  await expect(expr).toHaveText('SUM(A1:A2)+SQRT(D7)');

  // Step 1: SUM(A1:A2) → 7.
  await page.getByTestId('evaluate-formula-evaluate').click();
  await expect(expr).toHaveText('7+SQRT(D7)');

  // Step 2: SQRT(D7) → 3.
  await page.getByTestId('evaluate-formula-evaluate').click();
  await expect(expr).toHaveText('7+3');

  // Step 3: 7+3 → 10 (final).
  await page.getByTestId('evaluate-formula-evaluate').click();
  await expect(expr).toHaveText('10');
  await expect(page.getByTestId('evaluate-formula-evaluate')).toBeDisabled();

  // Restart returns to the original expression.
  await page.getByTestId('evaluate-formula-restart').click();
  await expect(expr).toHaveText('SUM(A1:A2)+SQRT(D7)');

  // The cell's formula was never overwritten by the off-cell evaluation.
  expect(await cellFormula(page, 'C1')).toBe('=SUM(A1:A2)+SQRT(D7)');
});

test('steps bare references individually before the final arithmetic', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);
  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('A1').setValue({ v: 10 });
    ws.getRange('B1').setValue({ v: 20 });
    ws.getRange('C1').setValue({ f: '=A1+B1*2' });
    ws.getRange('C1').activate();
  });

  await page.getByTestId('menubar-data').click();
  await page.getByTestId('menu-item-evaluate-formula').click();
  const expr = page.getByTestId('evaluate-formula-expr');
  await expect(expr).toHaveText('A1+B1*2');

  await page.getByTestId('evaluate-formula-evaluate').click();
  await expect(expr).toHaveText('10+B1*2'); // A1 stepped

  await page.getByTestId('evaluate-formula-evaluate').click();
  await expect(expr).toHaveText('10+20*2'); // B1 stepped

  await page.getByTestId('evaluate-formula-evaluate').click();
  await expect(expr).toHaveText('50'); // 10+20*2 final
});

test('shows a message when the active cell has no formula', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);
  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('A1').setValue({ v: 42 });
    ws.getRange('A1').activate();
  });

  await page.getByTestId('menubar-data').click();
  await page.getByTestId('menu-item-evaluate-formula').click();
  await expect(page.getByTestId('evaluate-formula-empty')).toBeVisible();
  await expect(page.getByTestId('evaluate-formula-evaluate')).toBeDisabled();
});
