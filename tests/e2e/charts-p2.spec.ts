import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';
import { mainCanvas, waitForUniver } from './_helpers';

/**
 * Charts P2 — Excel-style chart object interactions:
 *
 *   - Click selects (frame + 8 handles visible).
 *   - Click outside deselects.
 *   - Drag the body to move; cell anchor commits on release.
 *   - Drag any of the 8 handles to resize; cell anchor commits on release.
 *   - Delete / Backspace key removes the selected chart.
 *   - Esc clears the selection.
 *   - Right-click opens a context menu (Change type / Rename / Delete).
 */

async function seedAndInsert(page: Page) {
  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('A1').setValue({ v: 'R' });
    ws.getRange('B1').setValue({ v: 'Q1' });
    ws.getRange('C1').setValue({ v: 'Q2' });
    ws.getRange('A2').setValue({ v: 'N' });
    ws.getRange('B2').setValue({ v: 100 });
    ws.getRange('C2').setValue({ v: 120 });
    ws.getRange('A3').setValue({ v: 'S' });
    ws.getRange('B3').setValue({ v: 80 });
    ws.getRange('C3').setValue({ v: 95 });
    ws.getRange('A1:C3').activate();
  });
  await mainCanvas(page).first().click({ position: { x: 100, y: 100 } });
  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('A1:C3').activate();
  });
  await page.getByTestId('menubar-insert').click();
  await page.getByTestId('menu-item-insert-chart').click();
  await page.getByTestId('insert-chart-confirm').click();
  await expect(page.getByTestId('chart-overlay')).toBeVisible({ timeout: 5_000 });
}

test.describe('Charts P2 — Excel-style chart interactions', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await seedAndInsert(page);
  });

  test('click selects the chart and shows 8 resize handles', async ({ page }) => {
    const overlay = page.getByTestId('chart-overlay');
    await overlay.click();
    await expect(overlay).toHaveAttribute('data-selected', 'true');
    for (const h of ['n', 's', 'e', 'w', 'ne', 'nw', 'se', 'sw']) {
      await expect(page.getByTestId(`chart-handle-${h}`)).toBeVisible();
    }
  });

  test('clicking outside the chart deselects', async ({ page }) => {
    const overlay = page.getByTestId('chart-overlay');
    await overlay.click();
    await expect(overlay).toHaveAttribute('data-selected', 'true');
    // Click on a cell well above the chart's anchor (chart drops 2
    // rows below A1:C3, so it sits around row 4+; click row 1).
    await mainCanvas(page).first().click({ position: { x: 60, y: 30 } });
    await expect(overlay).not.toHaveAttribute('data-selected', 'true');
    await expect(page.getByTestId('chart-handle-nw')).toHaveCount(0);
  });

  test('Delete key removes the selected chart', async ({ page }) => {
    const overlay = page.getByTestId('chart-overlay');
    await overlay.click();
    await page.keyboard.press('Delete');
    await expect(page.getByTestId('chart-overlay')).toHaveCount(0);
  });

  test('Escape clears the chart selection without deleting', async ({ page }) => {
    const overlay = page.getByTestId('chart-overlay');
    await overlay.click();
    await page.keyboard.press('Escape');
    await expect(overlay).toBeVisible();
    await expect(overlay).not.toHaveAttribute('data-selected', 'true');
  });

  test('dragging the chart body moves the cell anchor', async ({ page }) => {
    const overlay = page.getByTestId('chart-overlay');
    const before = await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const evts = (window as any).__chartModelPos as Record<string, unknown> | undefined;
      return evts;
    });
    expect(before).toBeUndefined(); // sanity — we'll read pos via DOM later

    const box = await overlay.boundingBox();
    expect(box).not.toBeNull();
    const startX = box!.x + box!.width / 2;
    const startY = box!.y + box!.height / 2;
    const dropX = startX + 140;
    const dropY = startY + 90;

    // Drive the drag with explicit mousedown/move/up so React's pointer
    // handlers fire in order (page.mouse.drag elides intermediate moves).
    await page.mouse.move(startX, startY);
    await page.mouse.down();
    await page.mouse.move(startX + 70, startY + 45, { steps: 5 });
    await page.mouse.move(dropX, dropY, { steps: 5 });
    await page.mouse.up();

    // After release, the overlay's box should have moved roughly with
    // the drag (cell snap may round to the nearest cell boundary).
    await page.waitForTimeout(150);
    const after = await overlay.boundingBox();
    expect(after).not.toBeNull();
    expect(after!.x - box!.x).toBeGreaterThan(60); // approximately +140 px
    expect(after!.y - box!.y).toBeGreaterThan(40); // approximately +90 px
  });

  test('dragging the SE corner handle grows the chart', async ({ page }) => {
    const overlay = page.getByTestId('chart-overlay');
    await overlay.click();
    const handle = page.getByTestId('chart-handle-se');
    const handleBox = await handle.boundingBox();
    expect(handleBox).not.toBeNull();
    const startX = handleBox!.x + handleBox!.width / 2;
    const startY = handleBox!.y + handleBox!.height / 2;
    const before = await overlay.boundingBox();
    expect(before).not.toBeNull();

    await page.mouse.move(startX, startY);
    await page.mouse.down();
    await page.mouse.move(startX + 80, startY + 60, { steps: 4 });
    await page.mouse.move(startX + 160, startY + 120, { steps: 4 });
    await page.mouse.up();
    await page.waitForTimeout(150);

    const after = await overlay.boundingBox();
    expect(after).not.toBeNull();
    expect(after!.width).toBeGreaterThan(before!.width + 50);
    expect(after!.height).toBeGreaterThan(before!.height + 30);
  });

  test('right-click opens the context menu; delete from it removes the chart', async ({ page }) => {
    const overlay = page.getByTestId('chart-overlay');
    await overlay.click({ button: 'right' });
    await expect(page.getByTestId('chart-context-menu')).toBeVisible();
    await page.getByTestId('chart-context-delete').click();
    await expect(page.getByTestId('chart-overlay')).toHaveCount(0);
  });

  test('context-menu rename updates the chart title in the panel', async ({ page }) => {
    const overlay = page.getByTestId('chart-overlay');
    await overlay.click({ button: 'right' });
    await page.getByTestId('chart-context-rename').click();
    const input = page.getByTestId('chart-context-rename-input');
    await input.fill('Revenue chart');
    await input.press('Enter');

    await page.getByTestId('menubar-data').click();
    await page.getByTestId('menu-item-charts-panel').click();
    await expect(page.getByTestId('charts-panel').getByText('Revenue chart')).toBeVisible();
  });

  test('context-menu change-type switches a column chart to line', async ({ page }) => {
    const overlay = page.getByTestId('chart-overlay');
    await overlay.click({ button: 'right' });
    await page.getByTestId('chart-context-change-type').click();
    await expect(page.getByTestId('insert-chart-dialog')).toBeVisible();
    await page.getByTestId('insert-chart-type-line').click();
    await page.getByTestId('insert-chart-confirm').click();

    await page.getByTestId('menubar-data').click();
    await page.getByTestId('menu-item-charts-panel').click();
    await expect(page.getByTestId('charts-panel').getByText('Line')).toBeVisible();
  });
});
