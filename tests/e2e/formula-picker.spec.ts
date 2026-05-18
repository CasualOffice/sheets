import { expect, test, type Page } from '@playwright/test';
import { mainCanvas, selectRange, waitForUniver } from './_helpers';

/**
 * Excel-style range picker / pointer mode for the formula bar.
 *
 * Behavior: while a formula is being typed in the bar (draft starts
 * with `=`), clicking a sheet tab switches sheets WITHOUT committing,
 * and selecting a cell/range splices a `Sheet2!A1:B5` reference at
 * the caret. Enter commits on the origin cell + restores the origin
 * sheet. Esc cancels and restores the origin.
 *
 * These tests assert the data layer (cell value + active sheet on
 * commit) rather than canvas click positions — canvas pixel clicks
 * are flaky in headless. Programmatic `range.activate()` doesn't
 * trigger SelectionChanged (per FormulaBar's comment), so we simulate
 * canvas selection by dispatching a real click on the canvas where
 * needed. Where that's impractical we use Univer's internal command.
 */

async function addAndPopulate(page: Page, name: string, cells: Record<string, number | string>) {
  await page.evaluate(
    ({ name, cells }) => {
      const api = window.__univerAPI!;
      const wb = api.getActiveWorkbook()!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (wb as any).insertSheet();
      const sheets = wb.getSheets();
      const created = sheets[sheets.length - 1];
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (created as any).setName(name);
      for (const [a1, v] of Object.entries(cells)) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        (created as any).getRange(a1).setValue({ v });
      }
      // Switch back to Sheet1.
      const first = wb.getSheets().find((s) => s.getSheetName() === 'Sheet1');
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      if (first) (wb as any).setActiveSheet(first);
    },
    { name, cells },
  );
}

async function activeSheetName(page: Page): Promise<string> {
  return page.evaluate(() => {
    const api = window.__univerAPI!;
    return api.getActiveWorkbook()!.getActiveSheet()!.getSheetName();
  });
}

async function readSheetCell(page: Page, sheetName: string, a1: string) {
  return page.evaluate(
    ({ sheetName, a1 }) => {
      const api = window.__univerAPI!;
      const wb = api.getActiveWorkbook()!;
      const target = wb.getSheets().find((s) => s.getSheetName() === sheetName);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      return target ? (target as any).getRange(a1).getCellData() : null;
    },
    { sheetName, a1 },
  );
}

/** Simulate the canvas selection-change event the picker listens for.
 *  Programmatic `.activate()` goes through `set-selections` which the
 *  picker intentionally ignores; the real canvas click path emits the
 *  facade's `SelectionChanged` event. We can't easily fire that from
 *  outside, so we use `api.fireEvent` via the internal injector. */
async function fireCanvasSelection(page: Page, sheetName: string, a1: string) {
  await page.evaluate(
    ({ sheetName, a1 }) => {
      const api = window.__univerAPI!;
      const wb = api.getActiveWorkbook()!;
      const target = wb.getSheets().find((s) => s.getSheetName() === sheetName);
      if (!target) throw new Error(`sheet not found: ${sheetName}`);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (wb as any).setActiveSheet(target);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const range = (target as any).getRange(a1);
      range.activate();
      // Trigger the facade's SelectionChanged event with the right shape.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const fire = (api as any).fireEvent ?? (api as any).fire;
      const raw = range.getRange();
      fire?.call(api, 'SelectionChanged', {
        workbook: wb,
        worksheet: target,
        selections: [raw],
      });
    },
    { sheetName, a1 },
  );
}

test.describe('Formula range picker (pointer mode)', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await addAndPopulate(page, 'Numbers', { A1: 10, A2: 20, A3: 30 });
    await selectRange(page, 'B1');
  });

  test('clicking another sheet keeps formula edit alive (no auto-commit)', async ({ page }) => {
    const input = page.getByTestId('formula-input');
    await input.click();
    await input.fill('=');
    // Click the Numbers tab — should switch sheets but the formula
    // bar must still show `=` and stay in edit mode.
    const numbersTab = page.locator('.sheet-tab', { hasText: 'Numbers' });
    await numbersTab.click();
    await expect(input).toHaveValue('=');
    // B1 on Sheet1 must still be empty — no premature commit.
    expect(await readSheetCell(page, 'Sheet1', 'B1')).toBeNull();
  });

  test('selecting a range on another sheet inserts a cross-sheet ref', async ({ page }) => {
    const input = page.getByTestId('formula-input');
    await input.click();
    await input.fill('=SUM(');
    // Move caret to end so the picker splices the ref after `(`.
    await input.evaluate((el: HTMLInputElement) => el.setSelectionRange(5, 5));
    // Simulate canvas selection on Numbers!A1:A3.
    await fireCanvasSelection(page, 'Numbers', 'A1:A3');
    await expect(input).toHaveValue('=SUM(Numbers!A1:A3');
  });

  test('Enter commits formula on the ORIGIN cell after picking cross-sheet ref', async ({ page }) => {
    const input = page.getByTestId('formula-input');
    await input.click();
    await input.fill('=SUM(');
    await input.evaluate((el: HTMLInputElement) => el.setSelectionRange(5, 5));
    await fireCanvasSelection(page, 'Numbers', 'A1:A3');
    // Append closing paren and commit.
    await input.evaluate((el: HTMLInputElement) =>
      el.setSelectionRange(el.value.length, el.value.length),
    );
    await page.keyboard.type(')');
    await page.keyboard.press('Enter');

    // After commit, the value lives on Sheet1!B1 (origin), not on
    // Numbers (where the picker happened to be).
    const cd = await readSheetCell(page, 'Sheet1', 'B1');
    expect(cd?.f).toBe('=SUM(Numbers!A1:A3)');
    // Origin sheet is restored.
    expect(await activeSheetName(page)).toBe('Sheet1');
  });

  test('Escape cancels and restores origin sheet', async ({ page }) => {
    const input = page.getByTestId('formula-input');
    await input.click();
    await input.fill('=');
    await page.locator('.sheet-tab', { hasText: 'Numbers' }).click();
    expect(await activeSheetName(page)).toBe('Numbers');
    // From any element, Escape should reach the formula bar's revert.
    await input.focus();
    await input.press('Escape');
    expect(await activeSheetName(page)).toBe('Sheet1');
    // Draft cleared.
    await expect(input).not.toHaveValue('=');
    // B1 untouched.
    expect(await readSheetCell(page, 'Sheet1', 'B1')).toBeNull();
  });

  test('plain text edit (no leading =) still commits on cell change', async ({ page }) => {
    // Regression: the picker bypass applies ONLY to formulas. Plain
    // text in the formula bar should still auto-commit when the user
    // clicks another cell — this is the legacy NameBox-style behavior.
    const input = page.getByTestId('formula-input');
    await input.click();
    await input.fill('plain text');
    // Click another cell on the same sheet (not a tab).
    await mainCanvas(page).first().click({ position: { x: 200, y: 200 } });
    await page.waitForTimeout(150);
    const cd = await readSheetCell(page, 'Sheet1', 'B1');
    expect(cd?.v).toBe('plain text');
  });
});
