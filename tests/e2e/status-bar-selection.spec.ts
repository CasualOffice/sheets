import { expect, test } from '@playwright/test';
import { mainCanvas, selectRange, waitForUniver } from './_helpers';

/**
 * Polish #2 — selection feedback an Excel user expects:
 *   - Status bar Sum / Average / Count was already wired; this run
 *     also asserts Min / Max / Numerical Count.
 *   - Stats appear for a single non-empty cell (Excel shows Count=1).
 *   - Stats aggregate across multi-range selections (Ctrl+click adds
 *     disjoint regions; the bag is summed as one).
 *   - "3R × 2C" dimension badge sits next to the Name Box when the
 *     selection spans more than one cell.
 */
test.describe('Status bar + Name Box selection feedback', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('A1').setValue({ v: 10 });
      ws.getRange('A2').setValue({ v: 20 });
      ws.getRange('A3').setValue({ v: 30 });
      ws.getRange('B1').setValue({ v: 'text' });
      ws.getRange('B2').setValue({ v: 5 });
      ws.getRange('B3').setValue({ v: 15 });
    });
  });

  test('Sum / Average / Min / Max / Count populate for a multi-cell numeric selection', async ({ page }) => {
    await selectRange(page, 'A1:A3');
    const stats = page.getByTestId('sheet-tabs-stats');
    await expect(stats.getByTestId('stat-avg')).toContainText('Average: 20');
    await expect(stats.getByTestId('stat-count')).toContainText('Count: 3');
    await expect(stats.getByTestId('stat-min')).toContainText('Min: 10');
    await expect(stats.getByTestId('stat-max')).toContainText('Max: 30');
    await expect(stats.getByTestId('stat-sum')).toContainText('Sum: 60');
  });

  test('Numerical Count appears separately when the selection mixes text + numbers', async ({ page }) => {
    await selectRange(page, 'B1:B3');
    const stats = page.getByTestId('sheet-tabs-stats');
    // 3 cells total, only 2 numeric (B2=5, B3=15). Excel's status bar
    // shows both "Count" (all non-empty cells) and "Numerical Count"
    // (cells that participated in Sum/Min/Max/Avg).
    await expect(stats.getByTestId('stat-count')).toContainText('Count: 3');
    await expect(stats.getByTestId('stat-num-count')).toContainText('Numerical Count: 2');
    await expect(stats.getByTestId('stat-sum')).toContainText('Sum: 20');
  });

  test('single non-empty cell still shows Count=1', async ({ page }) => {
    await selectRange(page, 'A1');
    const stats = page.getByTestId('sheet-tabs-stats');
    await expect(stats.getByTestId('stat-count')).toContainText('Count: 1');
    await expect(stats.getByTestId('stat-sum')).toContainText('Sum: 10');
  });

  test('multi-range selection (Ctrl+click) aggregates across every range', async ({ page }) => {
    // Multi-range — construct two disjoint ranges and push them via
    // the SetSelectionsOperation that Univer's own Ctrl+click path
    // uses. The status bar's aggregate should bag both together.
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      const wb = api.getActiveWorkbook()!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = wb.getActiveSheet();
      // Helper: build the `primary` ISelectionCell that Univer expects
      // for the selection's "active" cell. For multi-range, only the
      // last selection gets a non-null primary (matches Univer's own
      // Ctrl+click behaviour — the active cell is the last picked).
      const primaryFor = (r: { startRow: number; startColumn: number }) => ({
        actualRow: r.startRow,
        actualColumn: r.startColumn,
        isMerged: false,
        isMergedMainCell: false,
        startRow: r.startRow,
        startColumn: r.startColumn,
        endRow: r.startRow,
        endColumn: r.startColumn,
        rangeType: 0,
      });
      const ranges = [
        { startRow: 0, endRow: 2, startColumn: 0, endColumn: 0 }, // A1:A3
        { startRow: 1, endRow: 2, startColumn: 1, endColumn: 1 }, // B2:B3
      ];
      api.executeCommand('sheet.operation.set-selections', {
        unitId: wb.getId(),
        subUnitId: ws.getSheetId(),
        selections: ranges.map((range, i) => ({
          range,
          primary: i === ranges.length - 1 ? primaryFor(range) : null,
          style: null,
        })),
      });
    });
    const stats = page.getByTestId('sheet-tabs-stats');
    // Combined: A1:A3 (10+20+30) + B2:B3 (5+15) = 80 across 5 cells.
    await expect(stats.getByTestId('stat-sum')).toContainText('Sum: 80', { timeout: 3_000 });
    await expect(stats.getByTestId('stat-count')).toContainText('Count: 5');
  });

  test('"3R × 2C" dimension badge appears when the selection spans more than one cell', async ({ page }) => {
    // Single cell — badge hidden.
    await selectRange(page, 'A1');
    await expect(page.getByTestId('sel-dimensions')).toHaveCount(0);
    // Multi-cell — badge shows rows × cols.
    await selectRange(page, 'A1:B3');
    const dims = page.getByTestId('sel-dimensions');
    await expect(dims).toBeVisible();
    await expect(dims).toContainText('3R × 2C');
  });

  test('stats disappear when the selection collapses to an empty single cell', async ({ page }) => {
    // Pick a cell well away from the seeded A1:B3 block so it's empty.
    await selectRange(page, 'F10');
    await expect(page.getByTestId('sheet-tabs-stats')).toHaveCount(0);
  });
});
