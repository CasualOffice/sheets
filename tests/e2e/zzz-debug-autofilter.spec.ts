import { test } from '@playwright/test';
import { mainCanvas, waitForUniver } from './_helpers';

test('debug autofilter with console capture', async ({ page }) => {
  const logs: string[] = [];
  page.on('console', (msg) => logs.push(`[${msg.type()}] ${msg.text()}`));
  page.on('pageerror', (err) => logs.push(`[pageerror] ${err.message}`));

  await page.goto('/');
  await waitForUniver(page);
  await mainCanvas(page).first().click({ position: { x: 200, y: 200 } });
  await page.evaluate(() => {
    const api = (window as any).__univerAPI!;
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('A1').setValue({ v: 'h1' });
    ws.getRange('A2').setValue({ v: 'a' });
    ws.getRange('A1:A2').activate();
  });
  await page.keyboard.press('Control+Shift+L');
  await page.waitForTimeout(3000);

  console.log('--- LOGS CAPTURED ---');
  for (const l of logs) {
    if (l.includes('[debug]') || l.includes('pageerror') || l.includes('error') || l.includes('Error')) {
      console.log(l);
    }
  }
});
