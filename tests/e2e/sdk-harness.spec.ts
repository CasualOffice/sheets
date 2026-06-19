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
    // Univer renders the grid onto a sized <canvas> a frame or two after onReady;
    // a non-zero-size canvas means the render engine + plugin graph constructed
    // with no redi throw. (Don't use .first()/toBeVisible — the formula UI adds
    // 0-size overlay canvases that aren't the grid.)
    await page.waitForFunction(
      () => Array.from(document.querySelectorAll('canvas')).some((c) => c.clientWidth > 0),
      null,
      { timeout: 30_000 },
    );
  });

  test('formula engine computes (=1+2 → 3)', async ({ page }) => {
    const result = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const api = (window as any).__sdkHarnessAPI;
      // `api.univer` is the FUniver escape hatch on CasualSheetsAPI.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.univer.getActiveWorkbook().getActiveSheet();
      ws.getRange(0, 0).setValue({ f: '=1+2' });
      // Main-thread compute is near-synchronous, but poll to be safe.
      for (let i = 0; i < 100; i++) {
        const v = ws.getRange(0, 0).getValue();
        if (v === 3 || v === '3') return v;
        await new Promise((r) => setTimeout(r, 100));
      }
      return ws.getRange(0, 0).getValue();
    });
    expect(Number(result)).toBe(3);
  });

  test('CasualSheetsAPI: snapshot round-trips through loadSnapshot', async ({ page }) => {
    const out = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const api = (window as any).__sdkHarnessAPI;
      // Write a value via the facade, snapshot, reload into a fresh unit,
      // and confirm the value survived the dispose/recreate round-trip.
      api.univer.getActiveWorkbook().getActiveSheet().getRange(0, 0).setValue('hello');
      const snap = api.getSnapshot();
      api.loadSnapshot(snap);
      for (let i = 0; i < 50; i++) {
        const v = api.univer.getActiveWorkbook().getActiveSheet().getRange(0, 0).getValue();
        if (v === 'hello') return { ok: true, v };
        await new Promise((r) => setTimeout(r, 100));
      }
      return {
        ok: false,
        v: api.univer.getActiveWorkbook().getActiveSheet().getRange(0, 0).getValue(),
      };
    });
    expect(out.ok).toBe(true);
  });

  test('onChange streams a debounced snapshot after an edit', async ({ page }) => {
    const out = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const w = window as any;
      const before = w.__sdkHarnessChangeCount ?? 0;
      w.__sdkHarnessAPI.univer
        .getActiveWorkbook()
        .getActiveSheet()
        .getRange(5, 5)
        .setValue('changed');
      // Default debounce is 400ms; wait past it, then read the captured snapshot.
      for (let i = 0; i < 30; i++) {
        if ((w.__sdkHarnessChangeCount ?? 0) > before) break;
        await new Promise((r) => setTimeout(r, 100));
      }
      const snap = w.__sdkHarnessLastSnapshot;
      const cell = snap?.sheets?.[Object.keys(snap.sheets)[0]]?.cellData?.[5]?.[5];
      return {
        fired: (w.__sdkHarnessChangeCount ?? 0) > before,
        value: cell?.v,
      };
    });
    expect(out.fired).toBe(true);
    expect(out.value).toBe('changed');
  });

  test('CasualSheetsAPI: getSelection returns the active range', async ({ page }) => {
    const sel = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const api = (window as any).__sdkHarnessAPI;
      api.univer.getActiveWorkbook().getActiveSheet().getRange(2, 3).activate();
      // Selection commands settle on the next frame.
      await new Promise((r) => setTimeout(r, 200));
      return api.getSelection();
    });
    expect(sel).not.toBeNull();
    expect(sel.range.startRow).toBe(2);
    expect(sel.range.startColumn).toBe(3);
    expect(typeof sel.sheetId).toBe('string');
  });
});
