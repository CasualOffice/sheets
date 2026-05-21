import { expect, test } from '@playwright/test';
import { waitForUniver, mainCanvas } from './_helpers';

/**
 * Copy/paste fidelity — verify the common cases that round-trip
 * through Univer's clipboard service. We invoke the canonical sheet
 * commands (`sheet.command.copy` / `sheet.command.paste`) directly
 * through the command bus rather than synthesizing keyboard events,
 * because Playwright's synthetic `Ctrl+V` doesn't carry a real
 * `ClipboardEvent.clipboardData` in headless mode and the clipboard
 * write/read API is asynchronous in chromium.
 *
 * The service-level call is what the keyboard handlers fan out to
 * anyway, so this gives realistic coverage of what survives. Adding
 * cases here is the right home for any "X is lost on copy/paste"
 * symptom users report.
 */

async function selectAndCopy(page: import('@playwright/test').Page, a1: string) {
  await page.evaluate(async (cell) => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange(cell).activate();
    await new Promise((r) => requestAnimationFrame(() => r(null)));
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    await (api as any).executeCommand('univer.command.copy');
  }, a1);
  // Give the clipboard write a tick to finish.
  await page.waitForTimeout(80);
}

async function selectAndPaste(page: import('@playwright/test').Page, a1: string) {
  await page.evaluate(async (cell) => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange(cell).activate();
    await new Promise((r) => requestAnimationFrame(() => r(null)));
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    await (api as any).executeCommand('univer.command.paste', { value: 'default-paste' });
  }, a1);
  // Paste pipeline does a clipboard read → HTML parse → mutation; let
  // it settle before assertions.
  await page.waitForTimeout(150);
}

async function selectAndPasteFormatting(page: import('@playwright/test').Page, a1: string) {
  await page.evaluate(async (cell) => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange(cell).activate();
    await new Promise((r) => requestAnimationFrame(() => r(null)));
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    await (api as any).executeCommand('sheet.command.paste-format');
  }, a1);
  await page.waitForTimeout(150);
}

test.describe('copy/paste fidelity round-trip', () => {
  test.use({ permissions: ['clipboard-read', 'clipboard-write'] });

  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await mainCanvas(page).first().click({ position: { x: 50, y: 50 } });
  });

  test('plain values round-trip', async ({ page }) => {
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('A1').setValue({ v: 'hello' });
      ws.getRange('B1').setValue({ v: 42 });
      ws.getRange('A2').setValue({ v: 'world' });
      ws.getRange('B2').setValue({ v: 3.14 });
    });
    await selectAndCopy(page, 'A1:B2');
    await selectAndPaste(page, 'D1');
    const after = await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      return {
        d1: ws.getRange('D1').getValue(),
        e1: ws.getRange('E1').getValue(),
        d2: ws.getRange('D2').getValue(),
        e2: ws.getRange('E2').getValue(),
      };
    });
    expect(after.d1).toBe('hello');
    expect(String(after.e1)).toBe('42');
    expect(after.d2).toBe('world');
    // Floating-point: paste might re-parse as number — check loosely.
    expect(parseFloat(String(after.e2))).toBeCloseTo(3.14, 2);
  });

  test('formulas survive copy/paste with the relative shift', async ({ page }) => {
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('A1').setValue({ v: 1 });
      ws.getRange('A2').setValue({ v: 2 });
      ws.getRange('A3').setValue({ f: '=A1+A2' });
    });
    await selectAndCopy(page, 'A3');
    await selectAndPaste(page, 'B3');
    const cd = await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      return ws.getRange('B3').getCellData();
    });
    // The pasted formula should reference the column-B equivalents
    // (B1+B2) — Univer's paste shifts relative refs by the paste
    // offset. We only assert the formula is present and references
    // B-column cells; an exact equality is brittle if Univer changes
    // its A1-to-canonical normalization.
    expect(typeof cd?.f).toBe('string');
    expect(cd.f.toUpperCase()).toContain('B1');
    expect(cd.f.toUpperCase()).toContain('B2');
  });

  test('bold + color + fill survive the round-trip', async ({ page }) => {
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('A1').setValue({
        v: 'styled',
        s: { bl: 1, cl: { rgb: '#1a73e8' }, bg: { rgb: '#fff59d' } },
      });
    });
    await selectAndCopy(page, 'A1');
    await selectAndPaste(page, 'C1');
    const style = await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const wb: any = api.getActiveWorkbook()!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = wb.getActiveSheet();
      const cd = ws.getRange('C1').getCellData();
      const sRef = cd?.s;
      return typeof sRef === 'string' ? wb.getWorkbook().getStyles().get(sRef) : sRef;
    });
    expect(style?.bl).toBe(1);
    // color values normalize to lowercase / rgb tuple; just check it's set.
    expect(style?.cl).toBeTruthy();
    expect(style?.bg).toBeTruthy();
  });

  test('number format survives the round-trip', async ({ page }) => {
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('A1').setValue({
        v: 1234.5,
        s: { n: { pattern: '0.00' } },
      });
    });
    await selectAndCopy(page, 'A1');
    await selectAndPaste(page, 'C1');
    const style = await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const wb: any = api.getActiveWorkbook()!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = wb.getActiveSheet();
      const cd = ws.getRange('C1').getCellData();
      const sRef = cd?.s;
      const s = typeof sRef === 'string' ? wb.getWorkbook().getStyles().get(sRef) : sRef;
      return s?.n;
    });
    // Univer's clipboard hook for number format encodes through the
    // numfmt plugin — verify the pattern survives. We only assert
    // it's present + non-empty (the exact string can normalize).
    expect(typeof style?.pattern === 'string' && style.pattern.length > 0).toBe(true);
  });

  test('paste formatting only keeps the target value intact', async ({ page }) => {
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('A1').setValue({
        v: 'source',
        s: { bl: 1, bg: { rgb: '#fff59d' } },
      });
      ws.getRange('B1').setValue({ v: 'keep me' });
    });
    await selectAndCopy(page, 'A1');
    await selectAndPasteFormatting(page, 'B1');
    const after = await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const wb: any = api.getActiveWorkbook()!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = wb.getActiveSheet();
      const cd = ws.getRange('B1').getCellData();
      const sRef = cd?.s;
      const style = typeof sRef === 'string' ? wb.getWorkbook().getStyles().get(sRef) : sRef;
      return { value: ws.getRange('B1').getValue(), style };
    });
    expect(after.value).toBe('keep me');
    expect(after.style?.bl).toBe(1);
    expect(after.style?.bg).toBeTruthy();
  });

  test('paste formatting only is exposed in the Edit menu', async ({ page }) => {
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('A1').setValue({
        v: 'source',
        s: { bl: 1, bg: { rgb: '#fff59d' } },
      });
      ws.getRange('B1').setValue({ v: 'keep me' });
    });
    await selectAndCopy(page, 'A1');
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('B1').activate();
    });
    await page.getByTestId('menubar-edit').click();
    await expect(page.getByTestId('menu-item-paste-format')).toBeVisible();
    await page.getByTestId('menu-item-paste-format').click();
    await page.waitForTimeout(150);
    const after = await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const wb: any = api.getActiveWorkbook()!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = wb.getActiveSheet();
      const cd = ws.getRange('B1').getCellData();
      const sRef = cd?.s;
      const style = typeof sRef === 'string' ? wb.getWorkbook().getStyles().get(sRef) : sRef;
      return { value: ws.getRange('B1').getValue(), style };
    });
    expect(after.value).toBe('keep me');
    expect(after.style?.bl).toBe(1);
    expect(after.style?.bg).toBeTruthy();
  });
});
