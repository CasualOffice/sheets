import { expect, test } from '@playwright/test';
import { mainCanvas, waitForUniver } from './_helpers';

/**
 * Excel-keyboard parity for the bindings we wired in MenuBar's global
 * keydown handler. Univer ships ~38 shortcuts of its own (arrow nav,
 * Ctrl+B/I/U, F2, etc.); this spec only exercises the ones we *add*
 * — gaps a returning Excel user would expect to find:
 *
 *   - Ctrl+N     new workbook
 *   - Ctrl+O     open file (asserts the file picker opens; nothing to pick)
 *   - Ctrl+PageDown / Ctrl+PageUp   sheet navigation
 *   - Shift+F11  new sheet
 *   - Ctrl+Home  jump to A1
 *   - Ctrl+End   jump to bottom-right of used range
 *   - Ctrl+;     today's date
 *
 * Ctrl+K (hyperlink) and Ctrl+H (replace) open Univer-internal UIs we
 * don't own; their bindings are smoke-checked in the menu spec.
 */
test.describe('Excel shortcuts', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    // Click the canvas so focus is outside any input — text-input
    // shortcuts (Ctrl+F etc.) explicitly skip when focused on inputs.
    await mainCanvas(page).first().click({ position: { x: 200, y: 200 } });
  });

  test('Ctrl+N replaces the workbook with a fresh one', async ({ page }) => {
    // Stamp the current workbook with a marker so we can detect the swap.
    const beforeId = await page.evaluate(() => window.__univerAPI!.getActiveWorkbook()!.getId());
    await page.keyboard.press('Control+n');
    await page.waitForFunction(
      (id) => window.__univerAPI?.getActiveWorkbook()?.getId() !== id,
      beforeId,
      { timeout: 5_000 },
    );
    expect(await page.evaluate(() => window.__univerAPI!.getActiveWorkbook()!.getId())).not.toBe(beforeId);
  });

  test('Ctrl+PageDown / Ctrl+PageUp navigate between sheets', async ({ page }) => {
    // Add a second sheet — Univer's insertSheet auto-switches focus
    // to the new sheet, so explicitly land on the FIRST sheet before
    // testing PageDown navigation.
    const firstId = await page.evaluate(() => {
      const api = window.__univerAPI!;
      const wb = api.getActiveWorkbook()!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const start = (wb.getActiveSheet() as any).getSheetId();
      wb.insertSheet();
      const sheets = wb.getSheets();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const first = sheets.find((s: any) => s.getSheetId() === start) ?? sheets[0];
      wb.setActiveSheet(first);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      return (first as any).getSheetId();
    });
    await page.keyboard.press('Control+PageDown');
    await page.waitForFunction(
      (start) => {
        const api = window.__univerAPI!;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        return (api.getActiveWorkbook()!.getActiveSheet() as any).getSheetId() !== start;
      },
      firstId,
      { timeout: 3_000 },
    );
    await page.keyboard.press('Control+PageUp');
    await page.waitForFunction(
      (start) => {
        const api = window.__univerAPI!;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        return (api.getActiveWorkbook()!.getActiveSheet() as any).getSheetId() === start;
      },
      firstId,
      { timeout: 3_000 },
    );
  });

  test('Shift+F11 inserts a new sheet', async ({ page }) => {
    const before = await page.evaluate(() => window.__univerAPI!.getActiveWorkbook()!.getSheets().length);
    await page.keyboard.press('Shift+F11');
    await page.waitForFunction(
      (n) => (window.__univerAPI!.getActiveWorkbook()!.getSheets().length ?? 0) > n,
      before,
      { timeout: 3_000 },
    );
  });

  test('Ctrl+Home jumps to A1', async ({ page }) => {
    // Move somewhere else first.
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('E5').activate();
    });
    await page.keyboard.press('Control+Home');
    await page.waitForFunction(
      () => {
        const api = window.__univerAPI!;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const ws: any = api.getActiveWorkbook()!.getActiveSheet();
        const r = ws.getActiveRange()?.getRange();
        return r?.startRow === 0 && r?.startColumn === 0;
      },
      null,
      { timeout: 3_000 },
    );
  });

  test('Ctrl+End jumps to the bottom-right of the used range', async ({ page }) => {
    // Plant a single cell so the used range has a known extent.
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('D7').setValue({ v: 'last' });
      ws.getRange('A1').activate();
    });
    await page.keyboard.press('Control+End');
    await page.waitForFunction(
      () => {
        const api = window.__univerAPI!;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const ws: any = api.getActiveWorkbook()!.getActiveSheet();
        const r = ws.getActiveRange()?.getRange();
        // D7 = row 6, col 3 (0-indexed).
        return r?.startRow === 6 && r?.startColumn === 3;
      },
      null,
      { timeout: 3_000 },
    );
  });

  test("Ctrl+; inserts today's date in the active cell", async ({ page }) => {
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('A1').activate();
    });
    await page.keyboard.press('Control+Semicolon');
    await page.waitForFunction(
      () => {
        const api = window.__univerAPI!;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const ws: any = api.getActiveWorkbook()!.getActiveSheet();
        const v = ws.getRange('A1').getValue();
        return typeof v === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(v);
      },
      null,
      { timeout: 3_000 },
    );
  });
});
