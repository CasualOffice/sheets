import { expect, test } from '@playwright/test';
import { waitForUniver } from './_helpers';

/**
 * DV Input Message — the data-validation panel's Advance options now expose
 * Input-Message fields (fork). This verifies they render, toggle, and accept
 * input. The on-hover rendering of the saved message is covered separately by
 * dv-input-message.spec.ts.
 */

test('the DV panel exposes Input Message fields under Advance options', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);
  await page.evaluate(async () => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    await (window as any).__ensurePlugin__?.('dv');
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('A1').activate();
    // Adds a rule on the selection AND opens the panel to its detail form.
    await api.executeCommand('data-validation.command.addRuleAndOpen');
  });

  // Expand "Advance options" → the input-message toggle is present.
  await page.getByText('Advance options').click();
  const toggle = page.getByText('Show input message when cell is selected');
  await expect(toggle).toBeVisible();

  // Enabling it reveals the title + text fields, which accept input.
  await toggle.click();
  await page.getByPlaceholder('Input message title').fill('Heads up');
  await page.getByPlaceholder('Input message text').fill('Pick a fruit from the list');
  await expect(page.getByPlaceholder('Input message title')).toHaveValue('Heads up');
  await expect(page.getByPlaceholder('Input message text')).toHaveValue(
    'Pick a fruit from the list',
  );
});
