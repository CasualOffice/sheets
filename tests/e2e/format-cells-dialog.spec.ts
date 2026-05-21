import { expect, test, type Page } from '@playwright/test';
import { selectRange, waitForUniver } from './_helpers';

async function readCellStyle(page: Page, a1: string) {
  return page.evaluate((cell) => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const wb: any = api.getActiveWorkbook()!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = wb.getActiveSheet();
    const data = ws.getRange(cell).getCellData();
    if (!data) return null;
    return typeof data.s === 'string' ? (wb.getWorkbook().getStyles().get(data.s) ?? null) : (data.s ?? null);
  }, a1);
}

test.describe('Format Cells dialog', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('A1').setValue({ v: 1234.5 });
      ws.getRange('A1').activate();
    });
    await selectRange(page, 'A1');
  });

  test('opens from Ctrl+1 and applies common formatting back to the active cell', async ({ page }) => {
    await page.keyboard.press('Control+1');
    await expect(page.getByTestId('format-cells-dialog')).toBeVisible();

    await page.getByTestId('format-cells-tab-font').click();
    await page.getByTestId('format-cells-bold').check();
    await page.getByTestId('format-cells-font-size').fill('14');

    await page.getByTestId('format-cells-tab-alignment').click();
    await page.getByTestId('format-cells-align').selectOption('center');
    await page.getByTestId('format-cells-wrap').check();

    await page.getByTestId('format-cells-tab-fill').click();
    await page.getByTestId('format-cells-fill-color').fill('#ffeb3b');

    await page.getByTestId('format-cells-tab-number').click();
    await page.getByTestId('format-cells-number-preset').selectOption('currency');

    await page.getByTestId('format-cells-apply').click();
    await expect(page.getByTestId('format-cells-dialog')).toHaveCount(0);

    const style = await readCellStyle(page, 'A1');
    expect(style?.bl).toBe(1);
    expect(style?.fs).toBe(14);
    expect(style?.ht).toBe(2);
    expect(style?.tb).toBe(3);
    expect(style?.bg?.rgb?.toLowerCase()).toBe('#ffeb3b');
    expect(style?.n?.pattern).toBe('"$"#,##0.00');
  });

  test('pre-populates from the active cell style', async ({ page }) => {
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('A1').setValue({
        v: 12,
        s: {
          bl: 1,
          fs: 16,
          ht: 2,
          bg: { rgb: '#ffeb3b' },
          n: { pattern: '0.00%' },
        },
      });
      ws.getRange('A1').activate();
    });
    await selectRange(page, 'A1');

    await page.keyboard.press('Control+1');
    await expect(page.getByTestId('format-cells-dialog')).toBeVisible();

    await page.getByTestId('format-cells-tab-font').click();
    await expect(page.getByTestId('format-cells-bold')).toBeChecked();
    await expect(page.getByTestId('format-cells-font-size')).toHaveValue('16');

    await page.getByTestId('format-cells-tab-alignment').click();
    await expect(page.getByTestId('format-cells-align')).toHaveValue('center');

    await page.getByTestId('format-cells-tab-fill').click();
    await expect(page.getByTestId('format-cells-fill-color')).toHaveValue('#ffeb3b');

    await page.getByTestId('format-cells-tab-number').click();
    await expect(page.getByTestId('format-cells-number-preset')).toHaveValue('percent');
    await expect(page.getByTestId('format-cells-number-pattern')).toHaveValue('0.00%');
  });
});
