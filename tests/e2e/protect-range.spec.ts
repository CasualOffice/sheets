import { expect, test } from '@playwright/test';
import { waitForUniver } from './_helpers';

/**
 * Range protection — Phase 4, T4.4 (slice 1).
 *
 * "Data → Protect range" locks the current selection while the rest of the
 * sheet stays editable (finer than the workbook read-only toggle). It drives
 * Univer's worksheet-permission facade, which owns the edit veto + locked-range
 * rendering. "Remove range protection" clears every rule on the sheet.
 *
 * Verifies through the real menu UI + the permission facade: protecting a
 * selection creates a rule that governs its cells; removing it clears them.
 * (The creating user keeps edit rights — only *other* peers are blocked — so we
 * assert the cell is under a protection rule, not the creator's own canEdit.)
 */
test('Data → Protect range locks the selection; Remove clears it', async ({ page }) => {
  test.setTimeout(60_000);
  await page.goto('/');
  await waitForUniver(page);

  // Select A1:B2.
  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (api.getActiveWorkbook()!.getActiveSheet() as any).getRange('A1:B2').activate();
  });

  // Protect via the menu.
  await page.getByTestId('menubar-data').click();
  await page.getByTestId('menu-item-protect-range').click();

  // The facade now reports a rule, and A1 is governed by a protection rule.
  await expect
    .poll(async () =>
      page.evaluate(async () => {
        const api = window.__univerAPI!;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const perm = (api.getActiveWorkbook()!.getActiveSheet() as any).getWorksheetPermission();
        const rules = await perm.listRangeProtectionRules();
        const ruleOnA1 = await perm.debugCellPermission(0, 0);
        return { count: rules.length, protectedA1: !!ruleOnA1 };
      }),
    )
    .toEqual({ count: 1, protectedA1: true });

  // Remove via the menu → editability restored.
  await page.getByTestId('menubar-data').click();
  await page.getByTestId('menu-item-remove-range-protection').click();

  await expect
    .poll(async () =>
      page.evaluate(async () => {
        const api = window.__univerAPI!;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const perm = (api.getActiveWorkbook()!.getActiveSheet() as any).getWorksheetPermission();
        const rules = await perm.listRangeProtectionRules();
        const ruleOnA1 = await perm.debugCellPermission(0, 0);
        return { count: rules.length, protectedA1: !!ruleOnA1 };
      }),
    )
    .toEqual({ count: 0, protectedA1: false });
});
