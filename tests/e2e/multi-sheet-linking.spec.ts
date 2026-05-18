import { expect, test } from '@playwright/test';
import { waitForUniver } from './_helpers';

/**
 * Cross-sheet references — `=Sheet2!A1` style. Univer's formula
 * engine handles these natively; this spec locks that in across the
 * three paths a user can take to make one:
 *
 *   1. Type the formula directly.
 *   2. Use the click-a-tab range picker that the formula bar already
 *      wires (formula-bar.spec covers the picker itself; we don't
 *      re-test it here).
 *   3. Round-trip via xlsx and confirm the formula text + cached
 *      value both survive.
 *
 * Also locks in the sheet-name autocomplete in the formula
 * suggestion dropdown — type `=Sh` and `Sheet1` / `Sheet2` show up
 * alongside SUMIF / SUMPRODUCT.
 */

test.describe('Cross-sheet references', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
  });

  test('=Sheet2!A1 evaluates to the referenced cell', async ({ page }) => {
    const result = await page.evaluate(async () => {
      const api = window.__univerAPI!;
      const wb = api.getActiveWorkbook()!;
      wb.insertSheet();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const [s1, s2] = wb.getSheets() as any[];
      s2.getRange('A1').setValue({ v: 42 });
      wb.setActiveSheet(s1);
      s1.getRange('A1').setValue({ f: `=${s2.getSheetName()}!A1` });
      // Formula compute runs in the worker — give it a beat.
      await new Promise((r) => setTimeout(r, 300));
      return s1.getRange('A1').getCellData();
    });
    expect(result?.f).toBe('=Sheet2!A1');
    expect(result?.v).toBe(42);
  });

  test('formula text + cached value survive xlsx round-trip', async ({ page }) => {
    await page.evaluate(async () => {
      const api = window.__univerAPI!;
      const wb = api.getActiveWorkbook()!;
      wb.insertSheet();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const [s1, s2] = wb.getSheets() as any[];
      s2.getRange('B3').setValue({ v: 7 });
      wb.setActiveSheet(s1);
      s1.getRange('A1').setValue({ f: '=Sheet2!B3' });
      await new Promise((r) => setTimeout(r, 300));
    });

    // Pull in the xlsx converter the same way xlsx-hyperlinks.spec does.
    await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const mod = await import(/* @vite-ignore */ '/src/xlsx/index.ts' as any);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (window as any).__xlsx = mod;
    });

    const after = await page.evaluate(async () => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const snap = (api.getActiveWorkbook()! as any).save();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const xlsx = (window as any).__xlsx!;
      const blob = await xlsx.workbookDataToXlsx(snap);
      const buf = await blob.arrayBuffer();
      const reloaded = await xlsx.xlsxToWorkbookData(buf);
      const firstSheetId = reloaded.sheetOrder[0];
      return reloaded.sheets[firstSheetId].cellData[0][0];
    });
    expect(after?.f).toBe('=Sheet2!B3');
    expect(after?.v).toBe(7);
  });

  test('formula-bar suggests sheet names when typing =Sh', async ({ page }) => {
    // Insert a second sheet so there's more than one name to match.
    await page.evaluate(() => {
      window.__univerAPI!.getActiveWorkbook()!.insertSheet();
    });

    const input = page.getByTestId('formula-input');
    await input.click();
    await input.focus();
    await input.fill('=Sh');
    // Suggestion list opens. Sheet1 + Sheet2 should both be there
    // tagged as kind=sheet (function suggestions also show; we just
    // assert the sheet entries are present).
    const sheet1 = page.getByTestId('formula-suggestion-Sheet1').filter({ has: page.locator('[data-kind="sheet"]') });
    const sheet2 = page.getByTestId('formula-suggestion-Sheet2').filter({ has: page.locator('[data-kind="sheet"]') });
    // The list items themselves carry data-kind="sheet"; the filter
    // above checks self, so use a direct selector instead.
    await expect(page.locator('[data-testid="formula-suggestion-Sheet1"][data-kind="sheet"]')).toBeVisible();
    await expect(page.locator('[data-testid="formula-suggestion-Sheet2"][data-kind="sheet"]')).toBeVisible();
    void sheet1;
    void sheet2;
  });

  test('clicking a sheet-name suggestion inserts `Name!`', async ({ page }) => {
    await page.evaluate(() => {
      window.__univerAPI!.getActiveWorkbook()!.insertSheet();
    });
    const input = page.getByTestId('formula-input');
    await input.click();
    await input.focus();
    await input.fill('=Sh');
    await page.locator('[data-testid="formula-suggestion-Sheet2"][data-kind="sheet"]').click();
    await expect(input).toHaveValue('=Sheet2!');
  });
});
