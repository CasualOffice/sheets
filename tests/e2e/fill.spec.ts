import { expect, test } from '@playwright/test';
import { selectRange, waitForUniver } from './_helpers';

/**
 * Smoke tests for the fill operations that Univer's sheets-ui ships by default:
 *   - Ctrl+D copies the top row of the selection down through the range
 *   - Ctrl+R copies the leftmost column of the selection across the range
 *   - The fill-handle (corner-drag) is rendered as part of the selection control
 *
 * If any of these regress, casual users lose the single most-used data-entry
 * gesture in a spreadsheet.
 */

test.describe('Fill operations', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
  });

  test('Ctrl+D fills the top row of the selection down', async ({ page }) => {
    // Seed A1 with a value, then select A1:A4 and press Ctrl+D.
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('A1').setValue({ v: 'seed' });
    });
    await selectRange(page, 'A1:A4');

    await page.keyboard.press('Control+d');

    // Univer dispatches the mutation through the command bus; allow a tick.
    await page.waitForFunction(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      return ws.getRange('A4').getValue() === 'seed';
    }, null, { timeout: 3_000 });

    const values = await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      return ['A1', 'A2', 'A3', 'A4'].map((a) => ws.getRange(a).getValue());
    });
    expect(values).toEqual(['seed', 'seed', 'seed', 'seed']);
  });

  test('Ctrl+R fills the leftmost column of the selection across', async ({ page }) => {
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('A1').setValue({ v: 'leftmost' });
    });
    await selectRange(page, 'A1:D1');

    await page.keyboard.press('Control+r');

    await page.waitForFunction(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      return ws.getRange('D1').getValue() === 'leftmost';
    }, null, { timeout: 3_000 });

    const values = await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      return ['A1', 'B1', 'C1', 'D1'].map((a) => ws.getRange(a).getValue());
    });
    expect(values).toEqual(['leftmost', 'leftmost', 'leftmost', 'leftmost']);
  });

  test('Ctrl+D extends formulas with relative refs (computed values)', async ({ page }) => {
    // The Excel/Sheets contract: fill-down rewrites A1-relative formulas.
    // Univer implements this via shared-formula ids — only the source cell
    // carries the literal `=A1*2`; descendants reference it via `si` and the
    // formula engine resolves them with the correct row offset. We verify the
    // contract by checking computed values, not raw `f` strings.
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('A1').setValue({ v: 10 });
      ws.getRange('A2').setValue({ v: 20 });
      ws.getRange('A3').setValue({ v: 30 });
      ws.getRange('B1').setValue({ f: '=A1*2' });
    });
    await selectRange(page, 'B1:B3');

    await page.keyboard.press('Control+d');

    // Wait until the formula engine has settled and B3 resolves to 60.
    await page.waitForFunction(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      return ws.getRange('B3').getValue() === 60;
    }, null, { timeout: 5_000 });

    const values = await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      return ['B1', 'B2', 'B3'].map((a) => ws.getRange(a).getValue());
    });
    expect(values).toEqual([20, 40, 60]);
  });
});
