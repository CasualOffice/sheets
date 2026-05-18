import { expect, test, type Page } from '@playwright/test';
import { selectRange, waitForUniver } from './_helpers';

/**
 * Cross-sheet formula references — the heart of any real Excel workflow
 * (lookup tables on one tab, summaries on another). These tests prove
 * the engine resolves cross-sheet refs end-to-end through the actual
 * formula bar entry path, not just the API.
 *
 * We poll for the computed value because Univer's formula engine runs
 * in a Web Worker (`apps/web/src/univer/formula-worker.ts`); the cell
 * write returns synchronously but the value lands a tick or two later.
 */

async function setSheetCell(
  page: Page,
  sheetName: string,
  a1: string,
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  data: any,
) {
  await page.evaluate(
    ({ sheetName, a1, data }) => {
      const api = window.__univerAPI!;
      const wb = api.getActiveWorkbook()!;
      const target = wb.getSheets().find((s) => s.getSheetName() === sheetName);
      if (!target) throw new Error(`sheet not found: ${sheetName}`);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (target as any).getRange(a1).setValue(data);
    },
    { sheetName, a1, data },
  );
}

async function addSheetNamed(page: Page, name: string) {
  await page.evaluate((name) => {
    const api = window.__univerAPI!;
    const wb = api.getActiveWorkbook()!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (wb as any).insertSheet();
    const sheets = wb.getSheets();
    const created = sheets[sheets.length - 1];
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (created as any).setName(name);
  }, name);
}

async function readSheetCell(
  page: Page,
  sheetName: string,
  a1: string,
) {
  return page.evaluate(
    ({ sheetName, a1 }) => {
      const api = window.__univerAPI!;
      const wb = api.getActiveWorkbook()!;
      const target = wb.getSheets().find((s) => s.getSheetName() === sheetName);
      if (!target) return null;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      return (target as any).getRange(a1).getCellData();
    },
    { sheetName, a1 },
  );
}

async function activateSheet(page: Page, name: string) {
  await page.evaluate((name) => {
    const api = window.__univerAPI!;
    const wb = api.getActiveWorkbook()!;
    const target = wb.getSheets().find((s) => s.getSheetName() === name);
    if (!target) throw new Error(`sheet not found: ${name}`);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (wb as any).setActiveSheet(target);
  }, name);
}

async function readComputed(
  page: Page,
  sheetName: string,
  a1: string,
) {
  // Worker calc finishes asynchronously. Retry until we either see a
  // value or hit a `#…` error sentinel.
  return await expect
    .poll(
      async () => {
        return page.evaluate(
          ({ sheetName, a1 }) => {
            const api = window.__univerAPI!;
            const wb = api.getActiveWorkbook()!;
            const target = wb.getSheets().find((s) => s.getSheetName() === sheetName);
            if (!target) return undefined;
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const cd = (target as any).getRange(a1).getCellData();
            return cd?.v;
          },
          { sheetName, a1 },
        );
      },
      // Generous timeout — the formula engine runs in a Web Worker
      // (formula-worker.ts) and under parallel-worker load the worker
      // can take a few seconds to settle the calc-engine init.
      { message: `waiting for computed value at ${sheetName}!${a1}`, timeout: 15_000 },
    )
    .not.toBeUndefined();
}

test.describe('Cross-sheet formula references', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
  });

  test('=Sheet2!A1 reads from another sheet', async ({ page }) => {
    await addSheetNamed(page, 'Sheet2');
    await setSheetCell(page, 'Sheet2', 'A1', { v: 42 });
    // insertSheet activates the new tab — switch back before commit.
    await activateSheet(page, 'Sheet1');

    await selectRange(page, 'B1');
    await page.getByTestId('formula-input').fill('=Sheet2!A1');
    await page.getByTestId('formula-input').press('Enter');

    await readComputed(page, 'Sheet1', 'B1');
    const cd = await readSheetCell(page, 'Sheet1', 'B1');
    expect(cd?.f).toBe('=Sheet2!A1');
    expect(Number(cd?.v)).toBe(42);
  });

  test('VLOOKUP across sheets resolves', async ({ page }) => {
    await addSheetNamed(page, 'Lookup');
    await setSheetCell(page, 'Lookup', 'A1', { v: 'AAA' });
    await setSheetCell(page, 'Lookup', 'B1', { v: 'Apple' });
    await setSheetCell(page, 'Lookup', 'A2', { v: 'BBB' });
    await setSheetCell(page, 'Lookup', 'B2', { v: 'Banana' });
    await setSheetCell(page, 'Lookup', 'A3', { v: 'CCC' });
    await setSheetCell(page, 'Lookup', 'B3', { v: 'Cherry' });

    await setSheetCell(page, 'Sheet1', 'A1', { v: 'BBB' });
    await activateSheet(page, 'Sheet1');
    await selectRange(page, 'B1');
    await page.getByTestId('formula-input').fill('=VLOOKUP(A1, Lookup!A1:B3, 2, FALSE)');
    await page.getByTestId('formula-input').press('Enter');

    await readComputed(page, 'Sheet1', 'B1');
    const cd = await readSheetCell(page, 'Sheet1', 'B1');
    expect(String(cd?.v)).toBe('Banana');
  });

  test('Sheet name containing a space requires single-quoting', async ({ page }) => {
    await addSheetNamed(page, 'My Data');
    await setSheetCell(page, 'My Data', 'A1', { v: 99 });
    await activateSheet(page, 'Sheet1');

    await selectRange(page, 'C1');
    await page.getByTestId('formula-input').fill("='My Data'!A1");
    await page.getByTestId('formula-input').press('Enter');

    await readComputed(page, 'Sheet1', 'C1');
    const cd = await readSheetCell(page, 'Sheet1', 'C1');
    expect(Number(cd?.v)).toBe(99);
  });

  test('INDIRECT can build a cross-sheet reference from a string', async ({ page }) => {
    await addSheetNamed(page, 'Other');
    await setSheetCell(page, 'Other', 'A1', { v: 7 });
    await activateSheet(page, 'Sheet1');

    await selectRange(page, 'D1');
    await page.getByTestId('formula-input').fill('=INDIRECT("Other!A1")');
    await page.getByTestId('formula-input').press('Enter');

    await readComputed(page, 'Sheet1', 'D1');
    const cd = await readSheetCell(page, 'Sheet1', 'D1');
    expect(Number(cd?.v)).toBe(7);
  });

  test('SUM across a cross-sheet range aggregates correctly', async ({ page }) => {
    await addSheetNamed(page, 'Numbers');
    await setSheetCell(page, 'Numbers', 'A1', { v: 10 });
    await setSheetCell(page, 'Numbers', 'A2', { v: 20 });
    await setSheetCell(page, 'Numbers', 'A3', { v: 30 });
    await activateSheet(page, 'Sheet1');

    await selectRange(page, 'E1');
    await page.getByTestId('formula-input').fill('=SUM(Numbers!A1:A3)');
    await page.getByTestId('formula-input').press('Enter');

    await readComputed(page, 'Sheet1', 'E1');
    const cd = await readSheetCell(page, 'Sheet1', 'E1');
    expect(Number(cd?.v)).toBe(60);
  });
});
