/* eslint-disable @typescript-eslint/no-explicit-any */
import { expect, test } from '@playwright/test';

test('insert pivot renders output + Fields panel picks up the model', async ({ page }) => {
  await page.goto('/sdk-harness?chrome=full');
  await page.waitForFunction(() => (window as any).__sdkHarnessReady === true, null, {
    timeout: 20000,
  });
  await expect(page.getByTestId('cs-menubar')).toBeVisible({ timeout: 20000 });
  await page.evaluate(async () => {
    const api = (window as any).__sdkHarnessAPI;
    const ws = api.univer.getActiveWorkbook().getActiveSheet();
    ws.getRange(0, 0).setValue('Cat');
    ws.getRange(0, 1).setValue('Val');
    const rows: Array<[string, number]> = [
      ['Fruit', 10],
      ['Veg', 20],
      ['Fruit', 5],
      ['Veg', 15],
    ];
    rows.forEach(([c, v], i) => {
      ws.getRange(i + 1, 0).setValue(c);
      ws.getRange(i + 1, 1).setValue(v);
    });
    ws.getRange('A1:B5').activate();
    await new Promise((r) => setTimeout(r, 150));
  });
  await page.getByTestId('cs-menu-insert').click();
  await page.getByTestId('cs-menuitem-insert-pivot').click();
  await expect(page.getByTestId('cs-insert-pivot-dialog')).toBeVisible({ timeout: 5000 });
  await page.getByTestId('cs-insert-pivot-dest-existing').check();
  await page.getByTestId('cs-insert-pivot-anchor').fill('D1');
  await page.getByTestId('cs-insert-pivot-create').click();
  await expect(page.getByTestId('cs-insert-pivot-dialog')).toHaveCount(0, { timeout: 5000 });
  const out = await page.evaluate(() => {
    const api = (window as any).__sdkHarnessAPI;
    const ws = api.univer.getActiveWorkbook().getActiveSheet();
    return { v1: ws.getRange(1, 4).getValue(), v2: ws.getRange(2, 4).getValue() };
  });
  expect([Number(out.v1), Number(out.v2)].sort((a, b) => a - b)).toEqual([15, 35]);
  await page.getByTestId('cs-panel-rail-pivot').click();
  await expect(page.getByTestId('cs-pivot-fields-panel')).toBeVisible({ timeout: 5000 });
  await expect(page.getByTestId('cs-pivot-fields-empty')).toHaveCount(0);
  await page.screenshot({ path: 'screenshots/pivot-sdk-e2e.png' });
});
