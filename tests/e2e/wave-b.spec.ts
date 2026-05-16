import { expect, test } from '@playwright/test';
import { selectRange, waitForUniver } from './_helpers';

/**
 * Wave B coverage — verifies the new ribbon buttons dispatch the right
 * Univer commands. We don't try to drive the resulting dialogs/panels
 * (those are Univer-internal and version-fragile); the assertion is that
 * the command event fires.
 */

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
  // Univer commands may dispatch async mutations; give them a moment.
  await page.waitForTimeout(400);
  return page.evaluate((frag) => {
    return (window.__cmdLog ?? []).some((id) => id.includes(frag));
  }, commandFragment);
}

test.describe('Wave B — Insert tab', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await selectRange(page, 'A1:C3');
    await page.getByTestId('ribbon-tab-insert').click();
  });

  test('Insert Table dispatches add-table', async ({ page }) => {
    const fired = await clickFiresCommand(page, 'ribbon-btn-insert-table', 'add-table');
    expect(fired).toBe(true);
  });

  test('Insert Comment dispatches show-comment-modal', async ({ page }) => {
    await selectRange(page, 'B2');
    const fired = await clickFiresCommand(page, 'ribbon-btn-insert-comment', 'comment-modal');
    expect(fired).toBe(true);
  });

  test('Insert Hyperlink dispatches insert-hyper-link', async ({ page }) => {
    const fired = await clickFiresCommand(page, 'ribbon-btn-insert-hyperlink', 'hyper-link');
    expect(fired).toBe(true);
  });
});

test.describe('Wave B — Data tab', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await selectRange(page, 'A1:A5');
    await page.getByTestId('ribbon-tab-data').click();
  });

  test('Data Validation opens the panel', async ({ page }) => {
    const fired = await clickFiresCommand(page, 'ribbon-btn-data-validation', 'validation');
    expect(fired).toBe(true);
  });

  test('Conditional Formatting opens the panel', async ({ page }) => {
    const fired = await clickFiresCommand(
      page,
      'ribbon-btn-conditional-formatting',
      'conditional.formatting',
    );
    expect(fired).toBe(true);
  });
});

test.describe('Wave B — View tab', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await page.getByTestId('ribbon-tab-view').click();
  });

  test('Freeze top row dispatches set-first-row-frozen', async ({ page }) => {
    const fired = await clickFiresCommand(
      page,
      'ribbon-btn-freeze-first-row',
      'set-first-row-frozen',
    );
    expect(fired).toBe(true);
  });

  test('Toggle Gridlines dispatches toggle-gridlines', async ({ page }) => {
    const fired = await clickFiresCommand(
      page,
      'ribbon-btn-toggle-gridlines',
      'toggle-gridlines',
    );
    expect(fired).toBe(true);
  });

  test('Zoom 100 dispatches set-zoom-ratio', async ({ page }) => {
    const fired = await clickFiresCommand(page, 'ribbon-btn-zoom-100', 'zoom-ratio');
    expect(fired).toBe(true);
  });
});

test.describe('Wave B — Status bar zoom + undo/redo moved to tabs strip', () => {
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
    // React's onChange listens for the synthesized change event; fill works.
    await page.getByTestId('statusbar-zoom-slider').fill('150');
    await page.waitForTimeout(400);
    const fired = await page.evaluate(() =>
      (window.__cmdLog ?? []).some((id) => id.includes('zoom-ratio')),
    );
    expect(fired).toBe(true);
  });

  test('Undo / Redo are in the sheet-tabs strip (not the titlebar)', async ({ page }) => {
    // Buttons live under sheet-tabs now.
    await expect(page.getByTestId('sheet-tabs').getByTestId('qat-undo')).toBeVisible();
    await expect(page.getByTestId('sheet-tabs').getByTestId('qat-redo')).toBeVisible();
    // And not under titlebar.
    await expect(page.getByTestId('titlebar').getByTestId('qat-undo')).toHaveCount(0);
  });
});
