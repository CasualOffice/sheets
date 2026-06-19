import { expect, test } from '@playwright/test';

/**
 * Exercises the SDK's `<CasualSheets>` editor directly via the dev-only
 * `/sdk-harness` route (apps/web/src/sdk-harness/SdkHarness.tsx). The app
 * normally renders its own `UniverSheet`, so this is the only coverage of the
 * published editor component. Verification surface for the SDK restructure.
 */

test.describe('SDK editor (CasualSheets) via /sdk-harness', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/sdk-harness');
    await page.waitForFunction(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      () => (window as any).__sdkHarnessReady === true,
      null,
      { timeout: 30_000 },
    );
  });

  test('boots and renders the grid (clean DI, no duplicate Univer)', async ({ page }) => {
    await expect(page.getByTestId('sdk-harness')).toBeVisible();
    // Univer renders the grid onto a <canvas> a frame or two after onReady; its
    // presence means the plugin graph constructed without a redi throw.
    await expect(page.locator('canvas').first()).toBeVisible({ timeout: 15_000 });
  });

  // Enabled by Batch 2b (wire the formula worker + RPC into CasualSheets). The
  // SDK editor currently ships with the formula engine disabled
  // (notExecuteFormula:true; the formula plugins are dropped to avoid the
  // IRPCChannelService DI throw), so =1+2 stays uncomputed until then.
  test.fixme('formula engine computes (=1+2 → 3)', async ({ page }) => {
    const result = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const api = (window as any).__sdkHarnessAPI;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook().getActiveSheet();
      ws.getRange(0, 0).setValue({ f: '=1+2' });
      // Formula compute runs through the worker — poll for the result.
      for (let i = 0; i < 100; i++) {
        const v = ws.getRange(0, 0).getValue();
        if (v === 3 || v === '3') return v;
        await new Promise((r) => setTimeout(r, 100));
      }
      return ws.getRange(0, 0).getValue();
    });
    expect(Number(result)).toBe(3);
  });
});
