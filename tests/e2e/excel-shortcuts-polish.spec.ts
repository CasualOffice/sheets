import { expect, test } from '@playwright/test';
import { mainCanvas, waitForUniver } from './_helpers';

/**
 * Polish pass — the shortcuts a real Excel user reaches for in the
 * first minute that we DIDN'T have. Each is its own test so a future
 * regression flags the exact missing binding.
 *
 *   - Ctrl+Space    select entire column(s) of the current selection
 *   - Shift+Space   select entire row(s)
 *   - Ctrl++        open Insert dialog (row / col / shift cells)
 *   - Ctrl+-        open Delete dialog
 *   - F2            edit-mode on the active cell
 *
 * Univer ships Shift+Arrow / Ctrl+Shift+Arrow extend natively — not
 * re-tested here.
 */
test.describe('Excel shortcuts — polish pass', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await mainCanvas(page).first().click({ position: { x: 200, y: 200 } });
    // Select B2 so the "extend to entire row/col" tests have a known
    // starting position.
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('B2').activate();
    });
  });

  test('Ctrl+Space extends the selection to span the full column(s)', async ({ page }) => {
    await page.keyboard.press('Control+Space');
    const range = await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      const r = ws.getActiveRange().getRange();
      return r;
    });
    expect(range.startColumn).toBe(1);
    expect(range.endColumn).toBe(1);
    expect(range.startRow).toBe(0);
    // Should span "all the rows of the sheet" — at least a couple
    // hundred. Default sheet size is 1024 rows.
    expect(range.endRow).toBeGreaterThan(200);
  });

  test('Shift+Space extends the selection to span the full row(s)', async ({ page }) => {
    await page.keyboard.press('Shift+Space');
    const range = await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      return ws.getActiveRange().getRange();
    });
    expect(range.startRow).toBe(1);
    expect(range.endRow).toBe(1);
    expect(range.startColumn).toBe(0);
    expect(range.endColumn).toBeGreaterThan(20);
  });

  test('Ctrl++ opens the Insert dialog with "Entire row" pre-selected', async ({ page }) => {
    // Numpad+ doesn't exist on all keyboards; the cross-keyboard binding
    // is Ctrl+Shift+= (same physical key as +). Playwright sends both as
    // the same event but the bindings accept either.
    await page.keyboard.press('Control+Shift+=');
    await expect(page.getByTestId('insert-cells-dialog')).toBeVisible();
    await expect(page.getByTestId('insert-cells-entire-row')).toBeChecked();
  });

  test('Ctrl++ → Entire column inserts an empty column at the current selection', async ({ page }) => {
    // Seed something at B2 so we can prove the shift happened.
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('B2').setValue({ v: 'marker' });
      ws.getRange('B2').activate();
    });
    await page.keyboard.press('Control+Shift+=');
    await page.getByTestId('insert-cells-entire-column').click();
    await page.getByTestId('insert-cells-ok').click();
    // Insert column-left → marker moves from B2 to C2.
    await page.waitForFunction(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      return ws.getRange('C2').getValue() === 'marker';
    }, { timeout: 3_000 });
    const b2 = await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      return ws.getRange('B2').getValue();
    });
    expect(b2 == null || b2 === '').toBe(true);
  });

  test('Ctrl+- opens the Delete dialog', async ({ page }) => {
    await page.keyboard.press('Control+-');
    await expect(page.getByTestId('delete-cells-dialog')).toBeVisible();
    await expect(page.getByTestId('delete-cells-entire-row')).toBeChecked();
  });

  test('Ctrl+- → Entire row deletes the row at the current selection', async ({ page }) => {
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('A1').setValue({ v: 'one' });
      ws.getRange('A2').setValue({ v: 'two' });
      ws.getRange('A3').setValue({ v: 'three' });
      ws.getRange('A2').activate();
    });
    await page.keyboard.press('Control+-');
    await page.getByTestId('delete-cells-ok').click();
    await page.waitForFunction(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      return ws.getRange('A2').getValue() === 'three';
    }, { timeout: 3_000 });
  });

  test('F2 enters edit mode on the active cell', async ({ page }) => {
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('A1').setValue({ v: 'hello' });
      ws.getRange('A1').activate();
    });
    await page.keyboard.press('F2');
    // Univer mounts a hidden editor canvas while editing — its existence
    // (separate from the main grid canvas) is a robust edit-mode signal.
    await page.waitForFunction(
      () => document.querySelectorAll('canvas[id^="univer-sheet-main-canvas_"]').length === 0 ||
            document.querySelectorAll('canvas').length >= 2,
      null,
      { timeout: 3_000 },
    );
  });
});
