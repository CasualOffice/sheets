import { expect, test } from '@playwright/test';
import { readCell, readStyle, selectRange, waitForUniver } from './_helpers';

/**
 * Coverage for the depth added in Phase 1.3: Insert / Formulas / Data tabs,
 * plus the new Home additions (font family/size, font/fill colors, wrap,
 * borders, vertical alignment).
 */

test.describe('Home — font, color, wrap, borders, vertical align', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await selectRange(page, 'A1');
  });

  test('Font family changes the cell style', async ({ page }) => {
    await page.getByTestId('ribbon-select-font-family').selectOption('Arial');
    const style = (await readStyle(page, 'A1')) as { ff?: string } | null;
    expect(style?.ff).toBe('Arial');
  });

  test('Font size changes the cell style', async ({ page }) => {
    await page.getByTestId('ribbon-select-font-size').selectOption('18');
    const style = (await readStyle(page, 'A1')) as { fs?: number } | null;
    expect(style?.fs).toBe(18);
  });

  test('Font color icon applies the default color', async ({ page }) => {
    await page.getByTestId('ribbon-color-font-color-apply').click();
    const style = (await readStyle(page, 'A1')) as { cl?: { rgb?: string } } | null;
    expect(style?.cl?.rgb?.toLowerCase()).toBe('#000000');
  });

  test('Fill color icon applies the default highlight color', async ({ page }) => {
    await page.getByTestId('ribbon-color-fill-color-apply').click();
    const style = (await readStyle(page, 'A1')) as { bg?: { rgb?: string } } | null;
    expect(style?.bg?.rgb?.toLowerCase()).toBe('#ffeb3b');
  });

  test('Wrap text toggles the wrap flag', async ({ page }) => {
    const wrap = page.getByTestId('ribbon-btn-wrap-text');
    await expect(wrap).toHaveAttribute('aria-pressed', 'false');
    await wrap.click();
    await expect(wrap).toHaveAttribute('aria-pressed', 'true');
  });

  test('Vertical alignment buttons are exclusive', async ({ page }) => {
    await page.getByTestId('ribbon-btn-align-middle').click();
    await expect(page.getByTestId('ribbon-btn-align-middle')).toHaveAttribute(
      'aria-pressed',
      'true',
    );
    await page.getByTestId('ribbon-btn-align-bottom').click();
    await expect(page.getByTestId('ribbon-btn-align-bottom')).toHaveAttribute(
      'aria-pressed',
      'true',
    );
    await expect(page.getByTestId('ribbon-btn-align-middle')).toHaveAttribute(
      'aria-pressed',
      'false',
    );
  });

  test('Borders dropdown — default click applies all borders', async ({ page }) => {
    await selectRange(page, 'A1');
    await page.getByTestId('ribbon-dropdown-borders-apply').click();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const style = (await readStyle(page, 'A1')) as any;
    expect(style?.bd).toBeTruthy();
    const hasAny = ['t', 'b', 'l', 'r'].some((k) => style.bd[k] !== undefined);
    expect(hasAny).toBe(true);
  });

  test('Borders dropdown — caret opens the popover with all options', async ({ page }) => {
    await selectRange(page, 'A1');
    await page.getByTestId('ribbon-dropdown-borders-caret').click();
    await expect(page.getByTestId('ribbon-dropdown-borders-popover')).toBeVisible();

    await page.getByTestId('ribbon-dropdown-borders-item-bottom').click();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const style = (await readStyle(page, 'A1')) as any;
    expect(style?.bd?.b).toBeTruthy();
  });
});

test.describe('Insert via menu bar', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
  });

  test('Insert → Row above shifts the existing value down', async ({ page }) => {
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('A1').setValue({ v: 'first' });
    });
    await selectRange(page, 'A1');

    await page.getByTestId('menubar-insert').click();
    await page.getByTestId('menu-item-insert-row-above').click();

    const a1 = await readCell(page, 'A1');
    const a2 = await readCell(page, 'A2');
    expect(a1?.v ?? null).toBeNull();
    expect(a2?.v).toBe('first');
  });

  test('Insert → Column right shifts content one column right', async ({ page }) => {
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('B1').setValue({ v: 'col-b' });
    });
    await selectRange(page, 'A1');

    await page.getByTestId('menubar-insert').click();
    await page.getByTestId('menu-item-insert-col-right').click();

    const b1 = await readCell(page, 'B1');
    const c1 = await readCell(page, 'C1');
    expect(b1?.v ?? null).toBeNull();
    expect(c1?.v).toBe('col-b');
  });

  test('Insert → New sheet appends to the workbook', async ({ page }) => {
    const before = await page.evaluate(() => {
      const api = window.__univerAPI!;
      return api.getActiveWorkbook()!.getSheets().length;
    });

    await page.getByTestId('menubar-insert').click();
    await page.getByTestId('menu-item-new-sheet').click();

    const after = await page.evaluate(() => {
      const api = window.__univerAPI!;
      return api.getActiveWorkbook()!.getSheets().length;
    });
    expect(after).toBe(before + 1);
  });
});

test.describe('Toolbar — AutoSum and sort', () => {
  test('AutoSum default lands SUM formula in cell below the selection', async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);

    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('A1').setValue({ v: 10 });
      ws.getRange('A2').setValue({ v: 20 });
      ws.getRange('A3').setValue({ v: 30 });
    });
    await selectRange(page, 'A1:A3');

    await page.getByTestId('ribbon-dropdown-auto-sum-apply').click();

    const a4 = await readCell(page, 'A4');
    expect(a4?.f).toBe('=SUM(A1:A3)');
  });

  test('Sort ascending reorders the selected column', async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);

    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('A1').setValue({ v: 3 });
      ws.getRange('A2').setValue({ v: 1 });
      ws.getRange('A3').setValue({ v: 2 });
    });
    await selectRange(page, 'A1:A3');
    await page.getByTestId('ribbon-btn-sort-asc').click();

    const values = await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      return [
        ws.getRange('A1').getCellData()?.v,
        ws.getRange('A2').getCellData()?.v,
        ws.getRange('A3').getCellData()?.v,
      ];
    });
    expect(values).toEqual([1, 2, 3]);
  });
});
