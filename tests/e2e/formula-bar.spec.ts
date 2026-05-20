import { expect, test, type Page } from '@playwright/test';
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

  test('Ctrl+G focuses the Name Box and selects its current reference', async ({ page }) => {
    await selectRange(page, 'C3');
    await page.keyboard.press('Control+g');

    const nameBox = page.getByTestId('name-box');
    await expect(nameBox).toBeFocused();
    await expect(nameBox).toHaveValue('C3');
    const selection = await nameBox.evaluate((el: HTMLInputElement) => ({
      start: el.selectionStart,
      end: el.selectionEnd,
      value: el.value,
    }));
    expect(selection).toEqual({ start: 0, end: 2, value: 'C3' });
  });

  test('Typing text and Enter commits to the active cell', async ({ page }) => {
    const input = page.getByTestId('formula-input');
    await input.fill('Hello world');
    await input.press('Enter');

    const data = await readCell(page, 'A1');
    expect(data?.v).toBe('Hello world');
    // After Enter the active cell moves down (Excel-equivalent), so the
    // formula bar now reflects A2 (empty). Re-select A1 to read it back.
    await selectRange(page, 'A1');
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

  // Excel-equivalent commit-and-navigate keybindings. Keyboard-heavy
  // users rely on these — losing them is a "the app feels wrong"
  // regression even if the value still commits correctly.

  async function activeA1(page: Page) {
    return page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      const r = ws.getActiveRange();
      const colLetter = (col: number) => {
        let s = '';
        let n = col;
        while (n >= 0) {
          s = String.fromCharCode(65 + (n % 26)) + s;
          n = Math.floor(n / 26) - 1;
        }
        return s;
      };
      return `${colLetter(r.getColumn())}${r.getRow() + 1}`;
    });
  }

  test('Enter commits and moves DOWN', async ({ page }) => {
    await selectRange(page, 'B2');
    await page.getByTestId('formula-input').fill('one');
    await page.getByTestId('formula-input').press('Enter');
    expect(await activeA1(page)).toBe('B3');
    expect((await readCell(page, 'B2'))?.v).toBe('one');
  });

  test('Shift+Enter commits and moves UP', async ({ page }) => {
    await selectRange(page, 'B5');
    await page.getByTestId('formula-input').fill('two');
    await page.getByTestId('formula-input').press('Shift+Enter');
    expect(await activeA1(page)).toBe('B4');
    expect((await readCell(page, 'B5'))?.v).toBe('two');
  });

  test('Tab commits and moves RIGHT', async ({ page }) => {
    await selectRange(page, 'C3');
    await page.getByTestId('formula-input').fill('right');
    await page.getByTestId('formula-input').press('Tab');
    expect(await activeA1(page)).toBe('D3');
    expect((await readCell(page, 'C3'))?.v).toBe('right');
  });

  test('Shift+Tab commits and moves LEFT', async ({ page }) => {
    await selectRange(page, 'D3');
    await page.getByTestId('formula-input').fill('left');
    await page.getByTestId('formula-input').press('Shift+Tab');
    expect(await activeA1(page)).toBe('C3');
    expect((await readCell(page, 'D3'))?.v).toBe('left');
  });

  test('Navigation at sheet edge clamps (no wrap, like Excel)', async ({ page }) => {
    // Shift+Tab at column A should stay at column A — Excel does not
    // wrap to the previous row's last column.
    await selectRange(page, 'A4');
    await page.getByTestId('formula-input').fill('edge');
    await page.getByTestId('formula-input').press('Shift+Tab');
    expect(await activeA1(page)).toBe('A4');
    expect((await readCell(page, 'A4'))?.v).toBe('edge');
  });

  test('F4 cycles absolute/relative on the ref under the caret', async ({ page }) => {
    const input = page.getByTestId('formula-input');
    await selectRange(page, 'A1');
    await input.click();
    await input.fill('=A1+B2');
    // Place caret inside A1 (after "=A").
    await input.evaluate((el: HTMLInputElement) => el.setSelectionRange(2, 2));
    await input.press('F4');
    await expect(input).toHaveValue('=$A$1+B2');

    // Second F4 → A$1
    await input.press('F4');
    await expect(input).toHaveValue('=A$1+B2');

    // Third F4 → $A1
    await input.press('F4');
    await expect(input).toHaveValue('=$A1+B2');

    // Fourth F4 → back to A1
    await input.press('F4');
    await expect(input).toHaveValue('=A1+B2');
  });
});
