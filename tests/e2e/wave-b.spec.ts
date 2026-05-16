import { expect, test } from '@playwright/test';
import { selectRange, waitForUniver } from './_helpers';

import type { Page } from '@playwright/test';

declare global {
  interface Window {
    __cmdLog?: string[];
  }
}

async function attachCmdLog(page: Page) {
  await page.evaluate(() => {
    if (window.__cmdLog) return;
    const log: string[] = [];
    window.__cmdLog = log;
    const api = window.__univerAPI!;
    api.addEvent(api.Event.CommandExecuted, (e) => {
      const id = (e as { id?: string }).id;
      if (id) log.push(id);
    });
  });
}

async function clickFiresCommand(
  page: Page,
  buttonTestId: string,
  commandFragment: string,
): Promise<boolean> {
  await attachCmdLog(page);
  await page.evaluate(() => {
    window.__cmdLog!.length = 0;
  });
  await page.getByTestId(buttonTestId).click();
  await page.waitForTimeout(400);
  return page.evaluate((frag) => {
    return (window.__cmdLog ?? []).some((id) => id.includes(frag));
  }, commandFragment);
}

async function clickMenuItemFiresCommand(
  page: Page,
  menuId: string,
  itemId: string,
  commandFragment: string,
): Promise<boolean> {
  await attachCmdLog(page);
  await page.evaluate(() => {
    window.__cmdLog!.length = 0;
  });
  await page.getByTestId(`menubar-${menuId}`).click();
  await page.getByTestId(`menu-item-${itemId}`).click();
  await page.waitForTimeout(400);
  return page.evaluate((frag) => {
    return (window.__cmdLog ?? []).some((id) => id.includes(frag));
  }, commandFragment);
}

test.describe('Insert menu', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await selectRange(page, 'A1:C3');
  });

  test('Insert Table dispatches add-table', async ({ page }) => {
    const fired = await clickMenuItemFiresCommand(page, 'insert', 'insert-table', 'add-table');
    expect(fired).toBe(true);
  });

  test('Insert Comment dispatches show-comment-modal', async ({ page }) => {
    await selectRange(page, 'B2');
    const fired = await clickMenuItemFiresCommand(page, 'insert', 'insert-comment', 'comment-modal');
    expect(fired).toBe(true);
  });

  test('Insert Hyperlink dispatches insert-hyper-link', async ({ page }) => {
    const fired = await clickMenuItemFiresCommand(page, 'insert', 'insert-link', 'hyper-link');
    expect(fired).toBe(true);
  });
});

test.describe('Data menu', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await selectRange(page, 'A1:A5');
  });

  test('Data Validation opens the panel', async ({ page }) => {
    const fired = await clickMenuItemFiresCommand(page, 'data', 'data-validation', 'validation');
    expect(fired).toBe(true);
  });

  test('Conditional Formatting opens the panel', async ({ page }) => {
    const fired = await clickMenuItemFiresCommand(
      page,
      'data',
      'conditional-formatting',
      'conditional.formatting',
    );
    expect(fired).toBe(true);
  });
});

test.describe('View menu', () => {
  test('Freeze top row dispatches set-first-row-frozen', async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    const fired = await clickMenuItemFiresCommand(
      page,
      'view',
      'freeze-row',
      'set-first-row-frozen',
    );
    expect(fired).toBe(true);
  });

  test('Toggle Gridlines dispatches toggle-gridlines', async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    const fired = await clickMenuItemFiresCommand(
      page,
      'view',
      'toggle-gridlines',
      'toggle-gridlines',
    );
    expect(fired).toBe(true);
  });
});

test.describe('Toolbar — zoom slider + borders dropdown + undo/redo placement', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
  });

  test('Status bar zoom slider exists and changes the ratio', async ({ page }) => {
    await expect(page.getByTestId('statusbar-zoom-slider')).toBeVisible();
    await expect(page.getByTestId('statusbar-zoom-label')).toHaveText('100%');

    await attachCmdLog(page);
    await page.evaluate(() => {
      window.__cmdLog!.length = 0;
    });
    await page.getByTestId('statusbar-zoom-slider').fill('150');
    await page.waitForTimeout(400);
    const fired = await page.evaluate(() =>
      (window.__cmdLog ?? []).some((id) => id.includes('zoom-ratio')),
    );
    expect(fired).toBe(true);
  });

  test('Borders popover closes on outside click (canvas)', async ({ page }) => {
    await page.getByTestId('ribbon-dropdown-borders-caret').click();
    await expect(page.getByTestId('ribbon-dropdown-borders-popover')).toBeVisible();
    const grid = page.locator('[id^="univer-sheet-main-canvas_"]');
    const box = await grid.boundingBox();
    if (box) {
      await page.mouse.click(box.x + 300, box.y + 200);
    }
    await expect(page.getByTestId('ribbon-dropdown-borders-popover')).toHaveCount(0);
  });

  test('Undo / Redo on the sheet-tabs strip work', async ({ page }) => {
    // Make a styled change we can undo.
    await selectRange(page, 'A1');
    await page.getByTestId('ribbon-btn-bold').click();
    await page.getByTestId('sheet-tabs').getByTestId('qat-undo').click();
    // The bold toggle should now reflect false again.
    await expect(page.getByTestId('ribbon-btn-bold')).toHaveAttribute('aria-pressed', 'false');
  });

  test('Toolbar Insert/Comment button dispatches show-comment-modal', async ({ page }) => {
    await selectRange(page, 'B2');
    const fired = await clickFiresCommand(page, 'ribbon-btn-insert-comment', 'comment-modal');
    expect(fired).toBe(true);
  });

  test('AutoSum split-button default dispatches a SUM formula', async ({ page }) => {
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('A1').setValue({ v: 10 });
      ws.getRange('A2').setValue({ v: 20 });
    });
    await selectRange(page, 'A1:A2');
    await page.getByTestId('ribbon-dropdown-auto-sum-apply').click();
    const a3 = await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      return ws.getRange('A3').getCellData();
    });
    expect((a3 as { f?: string }).f).toBe('=SUM(A1:A2)');
  });
});
