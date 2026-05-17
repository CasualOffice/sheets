import { expect, test } from '@playwright/test';
import { readCell, selectRange, waitForUniver } from './_helpers';

test.describe('Formula bar', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await selectRange(page, 'A1');
  });

  test('Name Box shows active cell reference and tracks selection', async ({ page }) => {
    await expect(page.getByTestId('name-box')).toHaveValue('A1');
    await selectRange(page, 'C3');
    await expect(page.getByTestId('name-box')).toHaveValue('C3');
  });

  test('Typing a reference into the Name Box navigates to it', async ({ page }) => {
    const nameBox = page.getByTestId('name-box');
    await nameBox.click();
    await nameBox.fill('D7');
    await nameBox.press('Enter');

    await expect(page.getByTestId('name-box')).toHaveValue('D7');
    const activeA1 = await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      const r = ws.getActiveRange();
      const a = (col: number) => String.fromCharCode(65 + col);
      return `${a(r.getColumn())}${r.getRow() + 1}`;
    });
    expect(activeA1).toBe('D7');
  });

  test('Name Box accepts a range and selects it', async ({ page }) => {
    const nameBox = page.getByTestId('name-box');
    await nameBox.click();
    await nameBox.fill('B2:C4');
    await nameBox.press('Enter');

    const sel = await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      const r = ws.getActiveRange();
      return {
        row: r.getRow(),
        col: r.getColumn(),
        w: r.getWidth(),
        h: r.getHeight(),
      };
    });
    expect(sel).toEqual({ row: 1, col: 1, w: 2, h: 3 });
  });

  test('Invalid reference in Name Box leaves selection unchanged', async ({ page }) => {
    await selectRange(page, 'A1');
    const nameBox = page.getByTestId('name-box');
    await nameBox.click();
    await nameBox.fill('not a ref');
    await nameBox.press('Enter');
    await expect(page.getByTestId('name-box')).toHaveValue('A1');
  });

  test('Typing text and Enter commits to the active cell', async ({ page }) => {
    const input = page.getByTestId('formula-input');
    await input.fill('Hello world');
    await input.press('Enter');

    const data = await readCell(page, 'A1');
    expect(data?.v).toBe('Hello world');
    await expect(input).toHaveValue('Hello world');
  });

  test('Numeric input is stored as a number (not text)', async ({ page }) => {
    await page.getByTestId('formula-input').fill('42');
    await page.getByTestId('formula-input').press('Enter');

    const data = await readCell(page, 'A1');
    expect(data?.v).toBe(42);
  });

  test('Formula =1+2 commits as a formula and computes', async ({ page }) => {
    const input = page.getByTestId('formula-input');
    await input.fill('=1+2');
    await input.press('Enter');

    // Cell carries the formula text and the computed value.
    const data = await readCell(page, 'A1');
    expect(data?.f).toBe('=1+2');

    // Re-selecting the cell shows the formula in the bar (not the result).
    await selectRange(page, 'A1');
    await expect(input).toHaveValue('=1+2');
  });

  test('Escape reverts an in-progress edit', async ({ page }) => {
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('A1').setValue({ v: 'original' });
    });
    await selectRange(page, 'A1');

    const input = page.getByTestId('formula-input');
    await expect(input).toHaveValue('original');
    await input.click();
    await input.fill('discarded');
    await input.press('Escape');

    await expect(input).toHaveValue('original');
    expect((await readCell(page, 'A1'))?.v).toBe('original');
  });

  test('Commit + Cancel buttons enable only while editing', async ({ page }) => {
    await expect(page.getByTestId('formula-commit')).toBeDisabled();
    await expect(page.getByTestId('formula-cancel')).toBeDisabled();

    await page.getByTestId('formula-input').fill('typing');
    await expect(page.getByTestId('formula-commit')).toBeEnabled();
    await expect(page.getByTestId('formula-cancel')).toBeEnabled();
  });

  test('Empty input clears the cell', async ({ page }) => {
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('A1').setValue({ v: 'something' });
    });
    await selectRange(page, 'A1');

    const input = page.getByTestId('formula-input');
    await input.fill('');
    await input.press('Enter');

    const data = await readCell(page, 'A1');
    expect(data?.v == null || data?.v === '').toBe(true);
  });
});
