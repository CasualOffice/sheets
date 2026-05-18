import { expect, test, type Page } from '@playwright/test';
import { mainCanvas, readCell, selectRange, waitForUniver } from './_helpers';

/**
 * Paste-from-external-Excel fidelity. The existing `copy-paste.spec.ts`
 * tests Univer's internal clipboard round-trip; this file specifically
 * exercises the `text/html` shape that Microsoft Excel writes to the
 * system clipboard. We dispatch a synthetic ClipboardEvent carrying
 * Excel-flavored HTML so we don't need OS clipboard access in CI.
 *
 * If any of these fail it means our paste path drops something a real
 * Excel user would see disappear when they Ctrl+V'd from the desktop
 * app — the most common bug class for Excel-to-web migrations.
 *
 * Some assertions are tagged `.fixme` rather than `.skip` so they show
 * up in the report as known gaps until the underlying handler lands.
 */

/**
 * Minimal HTML shape Excel writes for a 2x3 range with a header row,
 * formulas, and mixed styles. Stripped of MS Office's noisy `<!--...-->`
 * conditionals; Univer's parser looks for `<table>` regardless.
 */
const EXCEL_HTML_BASIC = `
<meta http-equiv="content-type" content="text/html; charset=utf-8">
<table border="0" cellpadding="0" cellspacing="0">
  <tr height="20">
    <td style="font-weight:bold;color:#1f497d">Code</td>
    <td style="font-weight:bold;color:#1f497d">Qty</td>
  </tr>
  <tr height="20">
    <td>AAA</td>
    <td>10</td>
  </tr>
  <tr height="20">
    <td>BBB</td>
    <td>20</td>
  </tr>
</table>
`;

const EXCEL_HTML_MERGED = `
<table border="1">
  <tr>
    <td colspan="2">merged-header</td>
  </tr>
  <tr>
    <td>x</td>
    <td>y</td>
  </tr>
</table>
`;

const EXCEL_HTML_FORMULA = `
<table>
  <tr>
    <td>1</td>
    <td>2</td>
    <td x:fmla="=SUM(A1:B1)">3</td>
  </tr>
</table>
`;

async function pasteHtml(page: Page, html: string) {
  // Focus the grid canvas so the paste event lands where the clipboard
  // service is listening. Univer attaches its keyboard/clipboard
  // listeners on the host element it manages.
  await mainCanvas(page).first().click({ position: { x: 50, y: 50 } });
  await page.waitForTimeout(50);

  await page.evaluate(async (html) => {
    const dt = new DataTransfer();
    dt.setData('text/html', html);
    dt.setData('text/plain', '');
    // Dispatch to the document — Univer hooks the global paste listener
    // (matches Excel/Sheets behavior where any focused grid cell receives).
    const ev = new ClipboardEvent('paste', {
      clipboardData: dt,
      bubbles: true,
      cancelable: true,
    });
    document.dispatchEvent(ev);
    // Some implementations also listen on the active element.
    document.activeElement?.dispatchEvent(ev);
    // Let the HTML→mutation pipeline settle.
    await new Promise((r) => setTimeout(r, 200));
  }, html);
}

test.describe('Paste from external Excel (text/html clipboard)', () => {
  test.use({ permissions: ['clipboard-read', 'clipboard-write'] });

  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await selectRange(page, 'A1');
  });

  // TODO(paste): the synthetic `new ClipboardEvent('paste', { clipboardData })`
  // we dispatch in `pasteHtml` doesn't trigger Univer's clipboard
  // service hooks in headless Chromium — Univer reads from the live
  // `navigator.clipboard` (async API) instead of from the event's
  // `clipboardData`. Properly exercising this requires writing the
  // HTML to the real system clipboard via permissions + a user-input
  // keypress, which Playwright's `page.keyboard` can't gate on
  // permission prompts in CI. Marked `.fixme` so the assertions stay
  // in tree and light up when the harness is upgraded (likely via
  // CDP `Input.insertText` + clipboard mocking).
  test.fixme('basic 2-column / 3-row table values land in the right cells', async ({ page }) => {
    await pasteHtml(page, EXCEL_HTML_BASIC);

    const a1 = await readCell(page, 'A1');
    const b1 = await readCell(page, 'B1');
    const a3 = await readCell(page, 'A3');
    const b3 = await readCell(page, 'B3');

    expect(String(a1?.v ?? '')).toBe('Code');
    expect(String(b1?.v ?? '')).toBe('Qty');
    expect(String(a3?.v ?? '')).toBe('BBB');
    expect(Number(b3?.v)).toBe(20);
  });

  test.fixme('header row keeps its bold style', async ({ page }) => {
    await pasteHtml(page, EXCEL_HTML_BASIC);

    const style = await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const wb: any = api.getActiveWorkbook()!;
      const ws = wb.getActiveSheet();
      const cd = ws.getRange('A1').getCellData();
      const sRef = cd?.s;
      return typeof sRef === 'string' ? wb.getWorkbook().getStyles().get(sRef) : sRef;
    });
    // Excel emits `font-weight:bold` → Univer should set `bl: 1`.
    expect(style?.bl).toBe(1);
  });

  test.fixme('merged-cell HTML produces a merged range', async ({ page }) => {
    await pasteHtml(page, EXCEL_HTML_MERGED);
    // The first row's two cells should be merged. Either getMergedRanges
    // shows it, or B1 reports `null` (merged-into A1).
    const merges = await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      return (ws as any).getMergedRanges?.() ?? [];
    });
    expect(merges.length).toBeGreaterThan(0);
  });

  test.fixme('Excel formula in x:fmla attribute is preserved', async ({ page }) => {
    await pasteHtml(page, EXCEL_HTML_FORMULA);
    const cd = await readCell(page, 'C1');
    // Excel's HTML carries the original formula in `x:fmla=` — Univer
    // must read that attribute to round-trip the formula. Without it,
    // the pasted cell holds the literal precomputed value (3) but loses
    // the formula, breaking downstream recalc.
    expect(typeof cd?.f).toBe('string');
    expect(String(cd?.f).toUpperCase()).toContain('SUM');
  });

  test('paste does not throw — no silent error in the console', async ({ page }) => {
    const errors: string[] = [];
    page.on('pageerror', (e) => errors.push(e.message));
    page.on('console', (msg) => {
      if (msg.type() === 'error') errors.push(msg.text());
    });
    await pasteHtml(page, EXCEL_HTML_BASIC);
    // Any uncaught exception from the paste pipeline is a regression.
    // We allow Univer's own benign warnings (e.g. "no clipboard data")
    // but not raw thrown errors.
    const fatal = errors.filter((e) => /Error|TypeError|Uncaught/.test(e));
    expect(fatal).toEqual([]);
  });
});
