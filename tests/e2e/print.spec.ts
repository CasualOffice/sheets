import { expect, test } from '@playwright/test';
import { waitForUniver } from './_helpers';

/**
 * Print plumbing — we can't drive the native print dialog in Playwright, but
 * we can verify that triggering print injects an iframe whose srcdoc contains
 * the active sheet's used range as an HTML table.
 */
test('File → Print injects an HTML render of the active sheet', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);

  await page.evaluate(() => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('A1').setValue({ v: 'Item' });
    ws.getRange('B1').setValue({ v: 'Qty' });
    ws.getRange('A2').setValue({ v: 'Apples' });
    ws.getRange('B2').setValue({ v: 12 });
  });

  // No-op window.print so the dialog doesn't actually fire under Playwright.
  // Defined on the page before we trigger print so the iframe inherits it.
  await page.addInitScript(() => {
    window.print = () => {
      /* no-op */
    };
  });
  // Also no-op print on every new iframe (since iframes get their own window).
  await page.evaluate(() => {
    const origAppend = document.body.appendChild.bind(document.body);
    document.body.appendChild = function <T extends Node>(node: T): T {
      const ret = origAppend(node);
      if (node instanceof HTMLIFrameElement) {
        node.addEventListener('load', () => {
          if (node.contentWindow) node.contentWindow.print = () => {};
        });
      }
      return ret;
    } as typeof document.body.appendChild;
  });

  await page.getByTestId('menubar-file').click();
  await page.getByTestId('menu-item-print').click();

  // The iframe is briefly attached. Grab its srcdoc and inspect.
  const srcdoc = await page.waitForFunction(
    () => {
      const f = document.querySelector('iframe');
      return f?.srcdoc ?? null;
    },
    null,
    { timeout: 3000 },
  );
  const html = (await srcdoc.jsonValue()) as string;
  expect(html).toContain('<table>');
  expect(html).toContain('Item');
  expect(html).toContain('Apples');
  expect(html).toContain('12');
});
