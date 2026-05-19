import { expect, test, type Page } from '@playwright/test';
import { selectRange, waitForUniver } from './_helpers';

/**
 * Phase 1.1 — Home ribbon dispatches real Univer commands.
 * We avoid relying on canvas pixels and instead read back the cell's
 * IStyleData via the exposed `window.__univerAPI` to assert state.
 */

type CellStyle = {
  bl?: 0 | 1;
  it?: 0 | 1;
  ul?: { s?: 0 | 1 } | null;
  ht?: number | null;
  n?: { pattern: string } | null;
};

async function setActiveCellValue(page: Page, a1: string, value: number | string) {
  await page.evaluate(
    ([cell, v]) => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange(cell).setValue({ v });
    },
    [a1, value] as const,
  );
}

async function getActiveCellStyle(page: Page): Promise<CellStyle> {
  return page.evaluate<CellStyle>(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const wb: any = api.getActiveWorkbook()!;
    const ws = wb.getActiveSheet();
    const range = ws.getActiveRange();
    const data = ws.getRange(range.getRow(), range.getColumn()).getCellData();
    if (!data) return {};
    if (typeof data.s === 'string') {
      const style = wb.getWorkbook().getStyles().get(data.s);
      return style ?? {};
    }
    return data.s ?? {};
  });
}

test.describe('Home ribbon → Univer commands', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await selectRange(page, 'A1');
  });

  test('Bold toggles fontWeight on the active cell', async ({ page }) => {
    const bold = page.getByTestId('ribbon-btn-bold');

    await expect(bold).toHaveAttribute('aria-pressed', 'false');
    await bold.click();
    await expect(bold).toHaveAttribute('aria-pressed', 'true');
    expect((await getActiveCellStyle(page)).bl).toBe(1);

    await bold.click();
    await expect(bold).toHaveAttribute('aria-pressed', 'false');
    expect((await getActiveCellStyle(page)).bl ?? 0).toBe(0);
  });

  test('Italic and Underline are independent toggles', async ({ page }) => {
    await page.getByTestId('ribbon-btn-italic').click();
    await page.getByTestId('ribbon-btn-underline').click();
    const style = await getActiveCellStyle(page);
    expect(style.it).toBe(1);
    expect(style.ul?.s).toBe(1);
    await expect(page.getByTestId('ribbon-btn-italic')).toHaveAttribute('aria-pressed', 'true');
    await expect(page.getByTestId('ribbon-btn-underline')).toHaveAttribute('aria-pressed', 'true');
  });

  test('Alignment buttons set horizontal alignment exclusively', async ({ page }) => {
    await page.getByTestId('ribbon-btn-align-center').click();
    await expect(page.getByTestId('ribbon-btn-align-center')).toHaveAttribute('aria-pressed', 'true');
    await expect(page.getByTestId('ribbon-btn-align-left')).toHaveAttribute('aria-pressed', 'false');

    await page.getByTestId('ribbon-btn-align-right').click();
    await expect(page.getByTestId('ribbon-btn-align-right')).toHaveAttribute('aria-pressed', 'true');
    await expect(page.getByTestId('ribbon-btn-align-center')).toHaveAttribute('aria-pressed', 'false');
  });

  test('Currency number format applies a $#,##0.00 pattern', async ({ page }) => {
    await setActiveCellValue(page, 'A1', 1234.5);
    await selectRange(page, 'A1');
    await page.getByTestId('ribbon-btn-numfmt-currency').click();

    await expect(page.getByTestId('ribbon-btn-numfmt-currency')).toHaveAttribute('aria-pressed', 'true');
    const style = await getActiveCellStyle(page);
    expect(style.n?.pattern).toBe('"$"#,##0.00');
  });

  test('Reflects style of newly-selected cells', async ({ page }) => {
    await page.getByTestId('ribbon-btn-bold').click();
    await expect(page.getByTestId('ribbon-btn-bold')).toHaveAttribute('aria-pressed', 'true');

    // Move selection to a different, unstyled cell.
    await selectRange(page, 'B2');
    await expect(page.getByTestId('ribbon-btn-bold')).toHaveAttribute('aria-pressed', 'false');

    // Move back — bold should reflect again.
    await selectRange(page, 'A1');
    await expect(page.getByTestId('ribbon-btn-bold')).toHaveAttribute('aria-pressed', 'true');
  });
});

test.describe('Status bar live stats', () => {
  test('shows Count/Sum/Avg for a multi-cell numeric selection', async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);

    // Seed three numeric cells.
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('A1').setValue({ v: 10 });
      ws.getRange('A2').setValue({ v: 20 });
      ws.getRange('A3').setValue({ v: 30 });
    });

    await selectRange(page, 'A1:A3');

    await expect(page.getByTestId('stat-count')).toHaveText('Count: 3');
    await expect(page.getByTestId('stat-sum')).toHaveText('Sum: 60');
    // Excel-style label is "Average:" (status bar). The polish pass
    // also added Min/Max — assert those here so a future regression
    // catches them.
    await expect(page.getByTestId('stat-avg')).toHaveText('Average: 20');
    await expect(page.getByTestId('stat-min')).toHaveText('Min: 10');
    await expect(page.getByTestId('stat-max')).toHaveText('Max: 30');
  });

  test('hides stats for a single empty cell selection', async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    // Far-away cell with no data — Excel shows nothing.
    await selectRange(page, 'F10');
    await expect(page.getByTestId('sheet-tabs-stats')).toHaveCount(0);
  });
});
