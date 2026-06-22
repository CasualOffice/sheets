import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';
import { mainCanvas, waitForUniver } from './_helpers';

/**
 * Combo charts + secondary (dual) value axis.
 *
 * Excel's "Format Data Series → Secondary Axis" + "Combo" chart type,
 * built on our custom ECharts config: a series flagged for the
 * secondary axis gets `yAxisIndex: 1` and the option emits a two-entry
 * `yAxis` array. A per-series kind override (Bars / Line) mixes bar +
 * line in one plot.
 *
 * The chart's live ECharts option is mirrored onto
 * `window.__casualChartOptions[chartId]` by ChartOverlay, so we can
 * assert on the resolved config directly.
 */

type ChartOptionWindow = Window & {
  __casualChartOptions?: Record<
    string,
    { yAxis?: unknown; series?: Array<{ name?: string; type?: string; yAxisIndex?: number }> }
  >;
};

async function seedTwoSeriesAndInsertChart(page: Page) {
  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('A1').setValue({ v: 'Quarter' });
    ws.getRange('B1').setValue({ v: 'Revenue' });
    ws.getRange('C1').setValue({ v: 'Margin %' });
    ws.getRange('A2').setValue({ v: 'Q1' });
    ws.getRange('B2').setValue({ v: 1000 });
    ws.getRange('C2').setValue({ v: 12 });
    ws.getRange('A3').setValue({ v: 'Q2' });
    ws.getRange('B3').setValue({ v: 2000 });
    ws.getRange('C3').setValue({ v: 18 });
    ws.getRange('A4').setValue({ v: 'Q3' });
    ws.getRange('B4').setValue({ v: 1500 });
    ws.getRange('C4').setValue({ v: 22 });
    ws.getRange('A1:C4').activate();
  });
  await mainCanvas(page)
    .first()
    .click({ position: { x: 120, y: 120 } });
  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('A1:C4').activate();
  });
  await page.getByTestId('menubar-insert').click();
  await page.getByTestId('menu-item-insert-chart').click();
  await page.getByTestId('insert-chart-confirm').click();
  await expect(page.getByTestId('chart-overlay')).toBeVisible({ timeout: 5_000 });
}

/** Read the single chart's id from the exposed option map. ChartOverlay
 *  mirrors its resolved option here on first render; under a cold dev
 *  server with parallel workers that render can lag, so wait generously. */
async function getChartId(page: Page): Promise<string> {
  return await page
    .waitForFunction(
      () => {
        const w = window as unknown as ChartOptionWindow;
        const ids = Object.keys(w.__casualChartOptions ?? {});
        return ids.length === 1 ? ids[0] : null;
      },
      null,
      { timeout: 20_000 },
    )
    .then((h) => h.jsonValue() as Promise<string>);
}

/** Open the chart's right-click menu, then the Format dialog. */
async function openFormatDialog(page: Page, chartId: string) {
  await page.evaluate((id) => {
    document.dispatchEvent(
      new CustomEvent('casual-chart-contextmenu', { detail: { id, x: 300, y: 300 } }),
    );
  }, chartId);
  await expect(page.getByTestId('chart-context-menu')).toBeVisible();
  await page.getByTestId('chart-context-format').click();
  await expect(page.getByTestId('format-chart-dialog')).toBeVisible();
}

test.describe('Charts — combo + secondary (dual) value axis', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await seedTwoSeriesAndInsertChart(page);
  });

  test('a single-axis chart has a non-array yAxis by default', async ({ page }) => {
    const id = await getChartId(page);
    const yAxisIsArray = await page.evaluate((cid) => {
      const w = window as unknown as ChartOptionWindow;
      return Array.isArray(w.__casualChartOptions?.[cid]?.yAxis);
    }, id);
    expect(yAxisIsArray).toBe(false);
  });

  test('assigning a series to the secondary axis emits two yAxis entries', async ({ page }) => {
    const id = await getChartId(page);
    await openFormatDialog(page, id);

    // The combo / axis section renders one row per series; series index
    // 1 is the second data column ("Margin %"). Put it on the secondary
    // axis and render it as a line (a textbook combo: revenue bars +
    // margin line on the right axis).
    await page.getByTestId('format-chart-secondary-axis-1').check();
    await page.getByTestId('format-chart-series-kind-1').selectOption('line');
    await page.getByTestId('format-chart-apply').click();
    await expect(page.getByTestId('format-chart-dialog')).toBeHidden();

    // The rebuilt option must now carry a two-entry yAxis and route the
    // second series to yAxisIndex 1, rendered as a line.
    await expect
      .poll(async () =>
        page.evaluate((cid) => {
          const w = window as unknown as ChartOptionWindow;
          const opt = w.__casualChartOptions?.[cid];
          if (!opt || !Array.isArray(opt.yAxis)) return null;
          const series = opt.series ?? [];
          return {
            yAxisCount: (opt.yAxis as unknown[]).length,
            secondary: series.find((s) => s.name === 'Margin %') ?? null,
            primary: series.find((s) => s.name === 'Revenue') ?? null,
          };
        }, id),
      )
      .toMatchObject({
        yAxisCount: 2,
        secondary: { type: 'line', yAxisIndex: 1 },
        primary: { type: 'bar', yAxisIndex: 0 },
      });
  });

  test('clearing the secondary axis collapses back to a single yAxis', async ({ page }) => {
    const id = await getChartId(page);

    await openFormatDialog(page, id);
    await page.getByTestId('format-chart-secondary-axis-1').check();
    await page.getByTestId('format-chart-apply').click();
    await expect(page.getByTestId('format-chart-dialog')).toBeHidden();
    await expect
      .poll(async () =>
        page.evaluate((cid) => {
          const w = window as unknown as ChartOptionWindow;
          return Array.isArray(w.__casualChartOptions?.[cid]?.yAxis);
        }, id),
      )
      .toBe(true);

    // Re-open and uncheck — the option must drop back to a single axis.
    await openFormatDialog(page, id);
    await page.getByTestId('format-chart-secondary-axis-1').uncheck();
    await page.getByTestId('format-chart-apply').click();
    await expect(page.getByTestId('format-chart-dialog')).toBeHidden();
    await expect
      .poll(async () =>
        page.evaluate((cid) => {
          const w = window as unknown as ChartOptionWindow;
          return Array.isArray(w.__casualChartOptions?.[cid]?.yAxis);
        }, id),
      )
      .toBe(false);
  });
});
