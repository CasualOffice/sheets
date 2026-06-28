import { expect, test } from '@playwright/test';
import { mainCanvas, waitForUniver } from './_helpers';

/**
 * Data Validation Input Message (fork: sheets-data-validation-ui). A DV rule
 * with `showInputMessage` + a prompt shows an informational popup when its cell
 * is pointed at — Excel's Input Message. Verifies the new fork controller is
 * wired and renders.
 */

test('a DV input message shows an info popup on hover', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);

  // Data-validation plugin is lazy-loaded — ensure it before using its facade.
  await page.evaluate(async () => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    await (window as any).__ensurePlugin__?.('dv');
  });

  await page.evaluate(async () => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const wb: any = api.getActiveWorkbook()!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = wb.getActiveSheet();
    // A valid value so the error alert doesn't pre-empt the input message.
    ws.getRange('A1').setValue({ v: 'apple' });
    await api.executeCommand('sheet.command.addDataValidation', {
      unitId: wb.getId(),
      subUnitId: ws.getSheetId(),
      rule: {
        uid: 'dv-input-msg',
        type: 'list',
        formula1: 'apple,banana,cherry',
        ranges: [{ startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 }],
        showInputMessage: true,
        promptTitle: 'Heads up',
        prompt: 'Pick a fruit from the list',
      },
    });
    ws.getRange('A1').activate();
  });

  // Hover A1 (a large top-left cell ≈ 92,33 px inside the canvas). A move away
  // and back makes Univer's hover service register the cell enter.
  const canvas = mainCanvas(page).first();
  await canvas.hover({ position: { x: 60, y: 30 } });
  await page.mouse.move(220, 220);
  await canvas.hover({ position: { x: 92, y: 33 } });
  await expect(page.getByText('Pick a fruit from the list')).toBeVisible({ timeout: 6000 });
});
