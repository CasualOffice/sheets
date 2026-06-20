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

  test('lazy plugins idle-load (conditional-formatting command registers)', async ({ page }) => {
    // lazyPlugins defaults on; idleLoadAll registers the feature plugins after
    // first paint. The CF "add-conditional-rule" command is a stable marker that
    // the conditional-formatting plugin actually loaded into the editor.
    const has = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const hasCommand = (window as any).__sdkHarnessHasCommand as
        | ((id: string) => boolean)
        | undefined;
      for (let i = 0; i < 60; i++) {
        if (hasCommand?.('sheet.command.add-conditional-rule')) return true;
        await new Promise((r) => setTimeout(r, 100));
      }
      return hasCommand?.('sheet.command.add-conditional-rule') ?? false;
    });
    expect(has).toBe(true);
  });

  test('appearance="dark" flips Univer dark mode + container class', async ({ page }) => {
    await page.goto('/sdk-harness?appearance=dark');
    await page.waitForFunction(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      () => (window as any).__sdkHarnessReady === true,
      null,
      { timeout: 30_000 },
    );
    const dark = await page.evaluate(async () => {
      for (let i = 0; i < 30; i++) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        if ((window as any).__sdkHarnessIsDark?.()) break;
        await new Promise((r) => setTimeout(r, 100));
      }
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      return (window as any).__sdkHarnessIsDark?.();
    });
    expect(dark).toBe(true);
    // We mirror the class onto the editor container; Univer's Workbench also
    // applies it to <html> (its dark CSS is page-global by design).
    await expect(page.locator('[data-testid="casual-sheets"].univer-dark')).toHaveCount(1);
  });

  test('default appearance is light (no dark mode, no dark class)', async ({ page }) => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const dark = await page.evaluate(() => (window as any).__sdkHarnessIsDark?.());
    expect(dark).toBe(false);
    await expect(page.locator('.univer-dark')).toHaveCount(0);
  });

  test('CasualSheetsAPI: setTheme flips dark mode imperatively', async ({ page }) => {
    // Default mount is light; drive dark via the API ref.
    const result = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const w = window as any;
      const before = w.__sdkHarnessIsDark?.();
      w.__sdkHarnessAPI.setTheme('dark');
      for (let i = 0; i < 20; i++) {
        if (w.__sdkHarnessIsDark?.() === true) break;
        await new Promise((r) => setTimeout(r, 50));
      }
      const afterDark = w.__sdkHarnessIsDark?.();
      w.__sdkHarnessAPI.setTheme('light');
      const afterLight = w.__sdkHarnessIsDark?.();
      return { before, afterDark, afterLight };
    });
    expect(result.before).toBe(false);
    expect(result.afterDark).toBe(true);
    expect(result.afterLight).toBe(false);
  });

  test('chrome="minimal" renders the toolbar and Bold dispatches', async ({ page }) => {
    await page.goto('/sdk-harness?chrome=minimal');
    await page.waitForFunction(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      () => (window as any).__sdkHarnessReady === true,
      null,
      { timeout: 30_000 },
    );
    await expect(page.getByTestId('casual-sheets-toolbar')).toBeVisible();
    await expect(page.locator('[data-action="bold"]')).toBeVisible();
    // Put a value in A1, select it, then click Bold via the chrome toolbar.
    await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const api = (window as any).__sdkHarnessAPI;
      const ws = api.univer.getActiveWorkbook().getActiveSheet();
      ws.getRange(0, 0).setValue('x');
      ws.getRange(0, 0).activate();
      await new Promise((r) => setTimeout(r, 150));
    });
    await page.locator('[data-action="bold"]').click();
    // Verify via the snapshot: A1's resolved style has bold (bl === 1).
    const isBold = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const api = (window as any).__sdkHarnessAPI;
      for (let i = 0; i < 20; i++) {
        const snap = api.getSnapshot();
        const sheet = snap?.sheets?.[Object.keys(snap.sheets)[0]];
        const cell = sheet?.cellData?.[0]?.[0];
        // style is either inline (cell.s as object) or a ref into snap.styles.
        const style = cell && (typeof cell.s === 'string' ? snap.styles?.[cell.s] : cell.s);
        if (style?.bl === 1) return true;
        await new Promise((r) => setTimeout(r, 100));
      }
      return false;
    });
    expect(isBold).toBe(true);
  });

  test('chrome defaults to none (no toolbar/formula bar)', async ({ page }) => {
    await expect(page.getByTestId('casual-sheets-toolbar')).toHaveCount(0);
    await expect(page.getByTestId('casual-sheets-formula-bar')).toHaveCount(0);
  });

  test('chrome formula bar: name box tracks selection + edit commits', async ({ page }) => {
    await page.goto('/sdk-harness?chrome=minimal');
    await page.waitForFunction(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      () => (window as any).__sdkHarnessReady === true,
      null,
      { timeout: 30_000 },
    );
    await expect(page.getByTestId('casual-sheets-formula-bar')).toBeVisible();
    // Select B2 (row 1, col 1) — the name box should show "B2".
    await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const api = (window as any).__sdkHarnessAPI;
      api.univer.getActiveWorkbook().getActiveSheet().getRange(1, 1).activate();
      await new Promise((r) => setTimeout(r, 200));
    });
    await expect(page.getByTestId('casual-sheets-name-box')).toHaveText('B2');
    // Type a formula into the bar and commit with Enter → B2 computes to 5.
    const input = page.getByTestId('casual-sheets-formula-input');
    await input.fill('=2+3');
    await input.press('Enter');
    const value = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const api = (window as any).__sdkHarnessAPI;
      for (let i = 0; i < 30; i++) {
        const v = api.univer.getActiveWorkbook().getActiveSheet().getRange(1, 1).getValue();
        if (v === 5 || v === '5') return v;
        await new Promise((r) => setTimeout(r, 100));
      }
      return api.univer.getActiveWorkbook().getActiveSheet().getRange(1, 1).getValue();
    });
    expect(Number(value)).toBe(5);
  });

  test('chrome status bar: selection stats (Average/Count/Sum)', async ({ page }) => {
    await page.goto('/sdk-harness?chrome=minimal');
    await page.waitForFunction(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      () => (window as any).__sdkHarnessReady === true,
      null,
      { timeout: 30_000 },
    );
    await expect(page.getByTestId('casual-sheets-status-bar')).toBeVisible();
    // Put 1,2,3 in A1:A3 and select the range → Sum 6, Count 3, Average 2.
    await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const api = (window as any).__sdkHarnessAPI;
      const ws = api.univer.getActiveWorkbook().getActiveSheet();
      ws.getRange(0, 0).setValue(1);
      ws.getRange(1, 0).setValue(2);
      ws.getRange(2, 0).setValue(3);
      ws.getRange('A1:A3').activate();
      await new Promise((r) => setTimeout(r, 250));
    });
    await expect(page.locator('[data-stat="sum"]')).toHaveText('Sum: 6');
    await expect(page.locator('[data-stat="count"]')).toHaveText('Count: 3');
    await expect(page.locator('[data-stat="average"]')).toHaveText('Average: 2');
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
