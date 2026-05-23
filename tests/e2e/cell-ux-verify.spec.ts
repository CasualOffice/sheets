import { expect, test } from '@playwright/test';
import { mainCanvas, waitForUniver } from './_helpers';

/**
 * Polish #3 — lock in Univer's stock cell-level affordances. Each of
 * these works out of the box in our build today; the tests exist so
 * that a future Univer-version bump, plugin-list change, or app-side
 * regression flags the loss within one CI run instead of in front of
 * a user.
 *
 * Univer ships rich context menus (Copy/Cut/Paste, Insert/Delete,
 * Merge, Freeze, Sort, Comment, Link, Note). The "Format Cells…"
 * multi-tab dialog Excel users reach for is *not* in this PR — it's
 * tracked separately because it's a substantial dialog of its own
 * (number / font / border / fill / protection tabs) and worth its
 * own scoped PR.
 */

const EXPECTED_CELL_ACTIONS = [
  'Copy',
  'Cut',
  'Paste',
  // Paste Special is added by our `context-menu-extensions.ts` —
  // verifies the custom command + menu-factory wiring actually
  // surfaces in Univer's QUICK group alongside Cut / Copy / Paste.
  'Paste Special',
  'Clear',
  'Merge cells',
  'Unmerge',
  'Insert',
  'Delete',
] as const;

test.describe('Cell-level UX — Univer affordances verified end-to-end', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
  });

  test('right-click on a cell opens a menu with the Excel-essentials', async ({ page }) => {
    const canvas = mainCanvas(page).first();
    await canvas.click({ position: { x: 80, y: 30 } });
    await canvas.click({ button: 'right', position: { x: 80, y: 30 } });
    const menu = page.locator('section.univer-popup').first();
    await expect(menu).toBeVisible({ timeout: 3_000 });
    const text = await menu.innerText();
    // Every action a returning Excel user expects on first right-click
    // must be present. If Univer drops any of them in an upgrade, this
    // assertion flags the regression.
    for (const action of EXPECTED_CELL_ACTIONS) {
      expect(text).toContain(action);
    }
  });

  test('auto-fit column command grows the column when content overflows', async ({ page }) => {
    // Seed A1 with a wide string + force a known narrow column width
    // so the "grew" check is deterministic across machines.
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      const wb = api.getActiveWorkbook()!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = wb.getActiveSheet();
      ws.getRange('A1').setValue({
        v: 'a much longer string than the column default fits — should expand on autofit',
      });
      ws.setColumnWidth(0, 40);
      ws.getRange('A1').activate();
    });
    const widthBefore = await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      return ws.getColumnWidth(0);
    });
    expect(widthBefore).toBe(40);
    // Univer renamed the command from set-worksheet-col-auto-width to
    // set-col-auto-width — the double-click-divider hook fires this id.
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      const wb = api.getActiveWorkbook()!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = wb.getActiveSheet();
      api.executeCommand('sheet.command.set-col-auto-width', {
        unitId: wb.getId(),
        subUnitId: ws.getSheetId(),
        ranges: [{ startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }],
      });
    });
    await page.waitForTimeout(150);
    const widthAfter = await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      return ws.getColumnWidth(0);
    });
    expect(widthAfter).toBeGreaterThan(widthBefore);
  });

  test('auto-fill command extends a numeric series down', async ({ page }) => {
    // Drag-fill from the corner dot is mouse-pixel-precise (sits on the
    // selection's lower-right corner); we dispatch the underlying
    // command directly — same code path the drag triggers — so any
    // regression in the autofill engine flags here.
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('A1').setValue({ v: 1 });
      ws.getRange('A2').setValue({ v: 2 });
      ws.getRange('A1:A2').activate();
    });
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      const wb = api.getActiveWorkbook()!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = wb.getActiveSheet();
      api.executeCommand('sheet.command.auto-fill', {
        unitId: wb.getId(),
        subUnitId: ws.getSheetId(),
        sourceRange: { startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 },
        targetRange: { startRow: 0, endRow: 5, startColumn: 0, endColumn: 0 },
      });
    });
    await page.waitForTimeout(150);
    const values = await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      return [
        ws.getRange('A3').getValue(),
        ws.getRange('A4').getValue(),
        ws.getRange('A5').getValue(),
        ws.getRange('A6').getValue(),
      ];
    });
    expect(values).toEqual([3, 4, 5, 6]);
  });

  test('right-click on a column header opens the column-header menu', async ({ page }) => {
    // Column headers sit in the first ~20 px of the grid. The exact
    // header height depends on the build's render config; (x=60, y=8)
    // lands squarely inside the "A" column letter for the default
    // skeleton.
    const canvas = mainCanvas(page).first();
    await canvas.click({ button: 'right', position: { x: 60, y: 8 } });
    const menu = page.locator('section.univer-popup').first();
    await expect(menu).toBeVisible({ timeout: 3_000 });
    const text = await menu.innerText();
    // The column-header menu should at least carry Insert / Delete
    // / Hide for the column. Univer ships these natively.
    expect(text.toLowerCase()).toMatch(/insert|delete|hide|width/);
  });

  test('right-click on a row header opens the row-header menu', async ({ page }) => {
    const canvas = mainCanvas(page).first();
    // Row 1's header letter sits in the first ~46 px column gutter.
    // (x=12, y=40) lands inside row 1's header rectangle.
    await canvas.click({ button: 'right', position: { x: 12, y: 40 } });
    const menu = page.locator('section.univer-popup').first();
    await expect(menu).toBeVisible({ timeout: 3_000 });
    const text = await menu.innerText();
    expect(text.toLowerCase()).toMatch(/insert|delete|hide|height/);
  });
});
