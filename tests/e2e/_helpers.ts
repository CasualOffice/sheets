import type { Page } from '@playwright/test';

/**
 * Returns the main grid canvas (the one users interact with).
 * Univer also renders a hidden formula-editor canvas; that one has no stable
 * id and we never want to target it from tests.
 */
export function mainCanvas(page: Page) {
  // Canvas id encodes the workbook unit id, which is now allocated dynamically
  // per workbook (so Open / New don't collide on the previous id).
  return page.locator('[id^="univer-sheet-main-canvas_"]');
}

export async function waitForUniver(page: Page) {
  await mainCanvas(page).waitFor({ timeout: 15_000 });
  await page.waitForFunction(() => Boolean(window.__univerAPI), null, { timeout: 5_000 });
}

export async function selectRange(page: Page, a1: string) {
  await page.evaluate((cell) => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange(cell).activate();
  }, a1);
}

export async function readCell(page: Page, a1: string) {
  return page.evaluate((cell) => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    return ws.getRange(cell).getCellData();
  }, a1);
}

/**
 * Reads the resolved IStyleData for a cell — handles both inline styles
 * (cellData.s is the object) and interned styles (cellData.s is an id into
 * the workbook's style table).
 */
export async function readStyle(page: Page, a1: string) {
  return page.evaluate((cell) => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const wb: any = api.getActiveWorkbook()!;
    const ws = wb.getActiveSheet();
    const data = ws.getRange(cell).getCellData();
    if (!data) return null;
    if (typeof data.s === 'string') {
      return wb.getWorkbook().getStyles().get(data.s) ?? null;
    }
    return data.s ?? null;
  }, a1);
}
