import { expect, test } from '@playwright/test';
import { readCell, selectRange, waitForUniver } from './_helpers';

/**
 * Pasting a multi-line formula into the formula bar should keep the
 * formula valid. The bar is a single-line input — newlines can't
 * render — but the onPaste handler normalizes newlines to spaces so
 * the formula engine still parses what's committed.
 *
 * Synthetic ClipboardEvent dispatch via Playwright + headless
 * Chromium reliably triggers React's onPaste handler when we attach
 * the listener directly to the input element. We don't assert the
 * default-paste browser behavior (single-line) — that path lets the
 * native input insertion happen and is hard to test in headless.
 */

test.describe('Formula bar multi-line paste', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await selectRange(page, 'A1');
  });

  test('multi-line formula paste normalizes newlines and commits as valid formula', async ({ page }) => {
    const input = page.getByTestId('formula-input');
    await input.click();
    await input.focus();
    // React's onPaste binding listens via the React event system —
    // our handler reads `e.clipboardData.getData('text/plain')`. A
    // synthetic ClipboardEvent with `clipboardData: new DataTransfer()`
    // routes correctly via React when dispatched on the focused input.
    await input.evaluate((el: HTMLInputElement) => {
      const dt = new DataTransfer();
      dt.setData('text/plain', '=IF(\n  1>0,\n  "yes",\n  "no"\n)');
      const ev = new ClipboardEvent('paste', {
        clipboardData: dt,
        bubbles: true,
        cancelable: true,
      });
      el.dispatchEvent(ev);
    });
    // After paste, draft should hold the formula with newlines replaced
    // by spaces.
    await expect(input).toHaveValue(/=IF\(\s+1>0,\s+"yes",\s+"no"\s+\)/);
    // Commit and verify the cell holds a valid IF formula that
    // evaluates to "yes".
    await input.press('Enter');
    await selectRange(page, 'A1');
    const cd = await readCell(page, 'A1');
    expect(typeof cd?.f).toBe('string');
    expect(cd.f).toContain('IF');
    expect(String(cd?.v ?? '')).toBe('yes');
  });
});
