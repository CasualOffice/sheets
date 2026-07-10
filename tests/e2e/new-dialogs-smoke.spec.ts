/**
 * Copyright 2026 Casual Office
 *
 * Licensed under the Apache License, Version 2.0 (the "License").
 */

/**
 * Smoke test for the 12 built-in dialogs added for sheet#295. Each menu item
 * used to open nothing (no built-in registered); this asserts every dialog now
 * MOUNTS from the SDK chrome menu without a runtime error. The real risk is a
 * Univer facade side-effect import not being registered by the host plugin set,
 * which would throw on open — so we fail on any pageerror captured while opening.
 *
 * Runs against the SDK chrome (`/sdk-harness?chrome=full`), which renders
 * `<CasualSheets>`'s own `cs-menu-*` menu bar — the surface embedders (dochub)
 * use. The standalone app renders `chrome="none"` and its own menu, so it can't
 * exercise these.
 */
import { expect, test, type Page } from '@playwright/test';

const MENUS = ['edit', 'insert', 'format', 'data', 'file', 'view'];

// kind (dialog root testid `cs-<kind>-dialog`) → menu-item id (`cs-menuitem-<id>`).
// Most menu items reuse the kind as their id; a few don't.
const DIALOGS: Array<{ kind: string; itemId?: string }> = [
  { kind: 'data-validation' },
  { kind: 'conditional-formatting' },
  { kind: 'custom-sort', itemId: 'sort-custom' },
  { kind: 'paste-special' },
  { kind: 'insert-function' },
  { kind: 'name-manager' },
  { kind: 'insert-cells', itemId: 'edit-insert-cells' },
  { kind: 'delete-cells', itemId: 'edit-delete-cells' },
  { kind: 'goal-seek' },
  { kind: 'insert-chart' },
  { kind: 'insert-sparkline' },
  { kind: 'insert-pivot' },
];

async function openFromMenu(page: Page, itemId: string): Promise<boolean> {
  for (const m of MENUS) {
    const btn = page.getByTestId(`cs-menu-${m}`);
    if ((await btn.count()) === 0) continue;
    await btn.click();
    const item = page.getByTestId(`cs-menuitem-${itemId}`);
    if (await item.isVisible().catch(() => false)) {
      await item.click();
      return true;
    }
    await page.keyboard.press('Escape');
  }
  return false;
}

test.describe('sheet#295 built-in dialogs — open smoke', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/sdk-harness?chrome=full');
    await page.waitForFunction(
      () => (window as unknown as { __sdkHarnessReady?: boolean }).__sdkHarnessReady === true,
      null,
      { timeout: 20_000 },
    );
    await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const api = (window as any).__sdkHarnessAPI;
      const ws = api.univer.getActiveWorkbook().getActiveSheet();
      ws.getRange(0, 0).setValue('Name');
      ws.getRange(0, 1).setValue('Score');
      ws.getRange(1, 0).setValue('Ann');
      ws.getRange(1, 1).setValue(12);
      ws.getRange(2, 0).setValue('Bo');
      ws.getRange(2, 1).setValue(7);
      ws.getRange('A1:B3').activate();
      await new Promise((r) => setTimeout(r, 150));
    });
  });

  for (const { kind, itemId } of DIALOGS) {
    test(`${kind} opens without a runtime error`, async ({ page }) => {
      const errors: string[] = [];
      page.on('pageerror', (e) => errors.push(e.message));

      const opened = await openFromMenu(page, itemId ?? kind);
      expect(opened, `menu item for ${kind} not found in any menu`).toBe(true);

      const dialog = page.getByTestId(`cs-${kind}-dialog`);
      await expect(dialog).toBeVisible({ timeout: 5_000 });

      expect(errors, `pageerror(s) on opening ${kind}: ${errors.join(' | ')}`).toHaveLength(0);

      if (kind === 'data-validation' || kind === 'conditional-formatting') {
        await page.screenshot({ path: `screenshots/dialog-${kind}.png` });
      }
    });
  }
});
