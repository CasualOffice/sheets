import { expect, test } from '@playwright/test';
import { waitForUniver } from './_helpers';

/**
 * Data Validation "Time" type (fork). Excel exposes Time as a distinct
 * Allow-type (Whole / Decimal / List / Date / Time / Text length / Custom).
 * Univer's core enum + cell-edit time-picker already existed but no validator
 * or panel view was registered, so the type never appeared. This verifies it
 * now shows in the panel's Type selector and is selectable.
 */

test('the DV panel offers "Time" as a validation type', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);
  await page.evaluate(async () => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    await (window as any).__ensurePlugin__?.('dv');
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('A1').activate();
    await api.executeCommand('data-validation.command.addRuleAndOpen');
  });

  // The first select in the rule panel is the "Type" selector. Open it.
  const typeSelect = page.locator('[data-u-comp="select"]').first();
  await expect(typeSelect).toBeVisible();
  await typeSelect.click();

  // "Time" is now a registered validator → it appears as an option. The type
  // selector renders options as radio items; this dev harness shows the raw
  // i18n key (the title resolves to "Time" once locales are bundled), so we
  // assert on the validator's title key — proof the Time validator registered.
  const timeOption = page.getByRole('menuitemradio', {
    name: 'sheets-data-validation.time.title',
    exact: true,
  });
  await expect(timeOption).toBeVisible();
  await timeOption.click();

  // Selecting it sticks: the Type selector now reflects the Time type.
  await expect(typeSelect).toContainText('sheets-data-validation.time.title');
});
