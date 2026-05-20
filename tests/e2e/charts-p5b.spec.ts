import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';
import ExcelJS from 'exceljs';
import { mainCanvas, waitForUniver } from './_helpers';

/**
 * Charts P5b — when our xlsx is opened in Excel, the chart should be
 * visible. We can't ship our live ECharts canvas inside the xlsx, but
 * we can embed a PNG snapshot anchored to the same cell rectangle.
 *
 * The live editable ChartModel still rides along via the
 * `__casual_sheets_charts__` resource so re-opening in casual-sheets
 * re-attaches the full interactive chart. The image is for foreign
 * readers (Excel / Numbers / LibreOffice).
 */

async function seedAndInsertChart(page: Page) {
  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('A1').setValue({ v: 'Quarter' });
    ws.getRange('B1').setValue({ v: 'Revenue' });
    ws.getRange('A2').setValue({ v: 'Q1' });
    ws.getRange('B2').setValue({ v: 100 });
    ws.getRange('A3').setValue({ v: 'Q2' });
    ws.getRange('B3').setValue({ v: 200 });
    ws.getRange('A4').setValue({ v: 'Q3' });
    ws.getRange('B4').setValue({ v: 150 });
    ws.getRange('A1:B4').activate();
  });
  await mainCanvas(page).first().click({ position: { x: 100, y: 100 } });
  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('A1:B4').activate();
  });
  await page.getByTestId('menubar-insert').click();
  await page.getByTestId('menu-item-insert-chart').click();
  await page.getByTestId('insert-chart-confirm').click();
  await expect(page.getByTestId('chart-overlay')).toBeVisible({ timeout: 5_000 });
}

test.describe('Charts P5b — chart image survives in exported xlsx', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await seedAndInsertChart(page);
  });

  test('exported xlsx contains an embedded image on the chart sheet', async ({ page }) => {
    // Trigger Save (Ctrl+S) and capture the download.
    const dl = page.waitForEvent('download', { timeout: 10_000 });
    await page.keyboard.press('Control+s');
    const file = await dl;
    const stream = await file.createReadStream();
    const chunks: Buffer[] = [];
    for await (const chunk of stream) {
      chunks.push(chunk as Buffer);
    }
    const bytes = Buffer.concat(chunks);

    // Parse with ExcelJS and check for an image on the active sheet.
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(bytes);
    // The first worksheet is where the chart was inserted.
    const ws = wb.worksheets[0];
    expect(ws, 'first worksheet present').toBeDefined();
    const images = ws.getImages();
    expect(images.length, 'at least one image embedded for the chart').toBeGreaterThan(0);

    // The image should be a PNG (extension on the workbook media entry).
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const media = (wb as any).media as Array<{ type?: string; extension?: string }> | undefined;
    expect(media?.some((m) => m?.type === 'image' && m?.extension === 'png')).toBe(true);
  });

  test('chart sidecar resource still ships so re-open is editable', async ({ page }) => {
    // The PNG is for foreign readers; the JSON sidecar is what makes
    // re-open in casual-sheets re-attach an editable chart. Both should
    // survive — this locks the "embed image" change against accidentally
    // dropping the sidecar.
    const dl = page.waitForEvent('download', { timeout: 10_000 });
    await page.keyboard.press('Control+s');
    const file = await dl;
    const stream = await file.createReadStream();
    const chunks: Buffer[] = [];
    for await (const chunk of stream) {
      chunks.push(chunk as Buffer);
    }
    const bytes = Buffer.concat(chunks);

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(bytes);
    // The sidecar lives in a hidden sheet named __casual_sheets_resources__.
    const sidecar = wb.worksheets.find((s) => s.name === '__casual_sheets_resources__');
    expect(sidecar, 'sidecar resources sheet present').toBeDefined();
    // The sidecar contents are JSON chunks; concatenate and look for the chart payload tag.
    let json = '';
    sidecar?.eachRow((row) => {
      const v = row.getCell(1).value;
      if (typeof v === 'string') json += v;
    });
    expect(json).toContain('__casual_sheets_charts__');
  });
});
