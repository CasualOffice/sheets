import { expect, test } from '@playwright/test';
import { mainCanvas, waitForUniver } from './_helpers';

/**
 * The specific flow a demo user hit that broke the experience:
 *
 *   1. Type `=` in a cell to start a formula.
 *   2. While the formula entry is live, click a different sheet tab.
 *   3. Select a range on the new sheet (click + shift-click on cells).
 *   4. Press Enter — the formula should reference `Sheet2!A1:A3` and
 *      commit back to the origin cell.
 *
 * The pre-fix code silently failed at step 3: SelectionChanged didn't
 * fire on real canvas clicks (only on programmatic selection updates),
 * and even when it did the picker appended every ref instead of
 * replacing the in-progress one. This spec locks both fixes in.
 *
 * Also covers VLOOKUP across sheets — the user mentioned VLOOKUP
 * "wasn't working"; the real blocker was probably step 1-3 above
 * (couldn't enter the formula), not VLOOKUP itself.
 */
test.describe('Cross-sheet formula picker — demo-user flow', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
  });

  test('start a formula, click another sheet, pick a range, commit — formula references the other sheet', async ({ page }) => {
    // Two sheets with known data on each. Capture the second sheet's
    // auto-assigned id so we can click its tab by testid.
    const lookupSheetId = await page.evaluate(() => {
      const api = window.__univerAPI!;
      const wb = api.getActiveWorkbook()!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws1: any = wb.getActiveSheet();
      ws1.getRange('A1').setValue({ v: 'header' });
      wb.insertSheet('Lookup');
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const sheets = wb.getSheets() as any[];
      const lookup = sheets[sheets.length - 1];
      lookup.getRange('A1').setValue({ v: 10 });
      lookup.getRange('A2').setValue({ v: 20 });
      lookup.getRange('A3').setValue({ v: 30 });
      const id = lookup.getSheetId();
      wb.setActiveSheet(sheets[0]);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (sheets[0] as any).getRange('B2').activate();
      return id as string;
    });

    // Start the formula via the formula bar — same code path Excel
    // users hit when they click the formula bar to edit.
    const fb = page.getByTestId('formula-input');
    await fb.click();
    await fb.fill('=SUM(');

    // Click the Lookup tab — formula entry must survive the sheet swap.
    await page.getByTestId(`sheet-tab-${lookupSheetId}`).click();

    // Real user clicks (not programmatic activate) so SelectionMoveEnd
    // fires — the picker's auto-injection path depends on it.
    const canvas = mainCanvas(page).first();
    await canvas.click({ position: { x: 60, y: 30 } });
    await canvas.click({ position: { x: 60, y: 70 }, modifiers: ['Shift'] });
    await expect(fb).toHaveValue(/^=SUM\(Lookup!A1:A3$/, { timeout: 3_000 });

    // Close the paren + commit. FormulaBar restores us to the origin.
    await fb.focus();
    await page.keyboard.press('End');
    await page.keyboard.type(')');
    await page.keyboard.press('Enter');

    // The formula landed on the original sheet (sheets[0], cell B2)
    // and evaluates to 10+20+30 = 60.
    await page.waitForFunction(
      () => {
        const api = window.__univerAPI!;
        const wb = api.getActiveWorkbook()!;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const sheets = wb.getSheets() as any[];
        const v = sheets[0].getRange('B2').getValue();
        return v === 60 || v === '60';
      },
      null,
      { timeout: 5_000 },
    );
  });

  test('VLOOKUP across sheets evaluates correctly', async ({ page }) => {
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      const wb = api.getActiveWorkbook()!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws1: any = wb.getActiveSheet();
      ws1.getRange('A1').setValue({ v: 'apple' });
      wb.insertSheet('Prices');
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const sheets = wb.getSheets() as any[];
      const prices = sheets[sheets.length - 1];
      prices.getRange('A1').setValue({ v: 'apple' });
      prices.getRange('B1').setValue({ v: 100 });
      prices.getRange('A2').setValue({ v: 'banana' });
      prices.getRange('B2').setValue({ v: 200 });
      wb.setActiveSheet(sheets[0]);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (sheets[0] as any).getRange('B1').setValue({ f: '=VLOOKUP(A1,Prices!A1:B2,2,FALSE)' });
    });

    await page.waitForFunction(
      () => {
        const api = window.__univerAPI!;
        const wb = api.getActiveWorkbook()!;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const ws: any = wb.getActiveSheet();
        const v = ws.getRange('B1').getValue();
        return v === 100 || v === '100';
      },
      null,
      { timeout: 5_000 },
    );
  });
});
