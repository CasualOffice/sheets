import { expect, test, type Page } from '@playwright/test';
import { waitForUniver } from './_helpers';

/**
 * Round-trip coverage for two pieces of xlsx fidelity that don't live on the
 * snapshot's `cellData` and so need side-channels:
 *
 *   1. Hyperlinks — handled by `@univerjs/sheets-hyper-link`'s HyperLinkModel.
 *      Import captures `{text, hyperlink}` cells into a __pendingHyperlinks
 *      array; on Open we wait for `onUniverSheetCreated` and replay them via
 *      AddHyperLinkCommand. Export reads HyperLinkModel.getUnit() back.
 *
 *   2. Univer plugin `resources` (tables, conditional formatting, data
 *      validation, comments, drawings, …) — stashed as JSON in a hidden
 *      `__casual_sheets_resources__` sheet that the importer unpacks and
 *      drops from sheetOrder.
 */

declare global {
  interface Window {
    __xlsx?: {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      xlsxToWorkbookData: (buf: ArrayBuffer) => Promise<any>;
      workbookDataToXlsx: (
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        data: any,
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        extras?: { hyperlinks?: Record<string, Array<{ row: number; column: number; payload: string; display?: string }>> },
      ) => Promise<Blob>;
    };
    __getHyperLinks__?: () => Array<{
      subUnitId: string;
      row: number;
      column: number;
      payload: string;
      display?: string;
    }>;
  }
}

async function exposeConverters(page: Page) {
  await page.evaluate(async () => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const mod = await import(/* @vite-ignore */ '/src/xlsx/index.ts' as any);
    window.__xlsx = mod;
  });
}

/**
 * Build an xlsx fixture with hyperlinks via the bundled exporter, write it to
 * /tmp, and return the absolute path. The file is what the filechooser will
 * receive.
 */
async function writeHyperlinkFixture(page: Page, fixturePath: string) {
  const bytes: number[] = await page.evaluate(async () => {
    const data = {
      id: 'fx-1',
      rev: 1,
      name: 'fixture',
      appVersion: '0.22.1',
      locale: 1,
      styles: {},
      sheetOrder: ['fx-s1'],
      sheets: {
        'fx-s1': {
          id: 'fx-s1',
          name: 'Links',
          cellData: {
            0: { 0: { v: 'Anthropic' }, 1: { v: 'GitHub' } },
            1: { 0: { v: 'plain' } },
          },
          rowCount: 1024,
          columnCount: 128,
        },
      },
    };
    const blob = await window.__xlsx!.workbookDataToXlsx(data, {
      hyperlinks: {
        'fx-s1': [
          { row: 0, column: 0, payload: 'https://anthropic.com', display: 'Anthropic' },
          { row: 0, column: 1, payload: 'https://github.com/schnsrw/sheets', display: 'GitHub' },
        ],
      },
    });
    return Array.from(new Uint8Array(await blob.arrayBuffer()));
  });
  const fs = await import('node:fs');
  fs.writeFileSync(fixturePath, Buffer.from(bytes));
}

test.describe('xlsx hyperlinks & resources round-trip', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await exposeConverters(page);
  });

  test('importer captures cell hyperlinks into __pendingHyperlinks', async ({ page }) => {
    const pending = await page.evaluate(async () => {
      const original = {
        id: 'wb-1',
        rev: 1,
        name: 'wb',
        appVersion: '0.22.1',
        locale: 1,
        styles: {},
        sheetOrder: ['s1'],
        sheets: {
          s1: {
            id: 's1',
            name: 'Sheet1',
            cellData: { 0: { 0: { v: 'Anthropic' } } },
            rowCount: 1024,
            columnCount: 128,
          },
        },
      };
      const blob = await window.__xlsx!.workbookDataToXlsx(original, {
        hyperlinks: {
          s1: [{ row: 0, column: 0, payload: 'https://anthropic.com', display: 'Anthropic' }],
        },
      });
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const reloaded: any = await window.__xlsx!.xlsxToWorkbookData(await blob.arrayBuffer());
      return reloaded.__pendingHyperlinks ?? null;
    });

    expect(pending).not.toBeNull();
    expect(pending).toHaveLength(1);
    const [link] = pending!;
    expect(link.row).toBe(0);
    expect(link.column).toBe(0);
    expect(link.payload).toBe('https://anthropic.com');
    expect(link.display).toBe('Anthropic');
    expect(typeof link.id).toBe('string');
    expect(link.id.length).toBeGreaterThan(0);
    // sheetId on the side-channel must match what the importer produced for
    // this worksheet, so the replay step targets the right subUnit.
    expect(typeof link.subUnitId).toBe('string');
  });

  test('resources blob round-trips via the hidden sheet', async ({ page }) => {
    const result = await page.evaluate(async () => {
      const original = {
        id: 'wb-2',
        rev: 1,
        name: 'wb',
        appVersion: '0.22.1',
        locale: 1,
        styles: {},
        sheetOrder: ['s1'],
        sheets: {
          s1: {
            id: 's1',
            name: 'Sheet1',
            cellData: { 0: { 0: { v: 'hi' } } },
            rowCount: 1024,
            columnCount: 128,
          },
        },
        // Simulate a real plugin resource — exactly the shape Univer uses.
        resources: [
          { name: 'SHEET_TABLE_PLUGIN', data: JSON.stringify({ tables: [{ id: 't1' }] }) },
          { name: 'SHEET_DATA_VALIDATION_PLUGIN', data: '{}' },
        ],
      };
      const blob = await window.__xlsx!.workbookDataToXlsx(original);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const reloaded: any = await window.__xlsx!.xlsxToWorkbookData(await blob.arrayBuffer());
      return {
        resources: reloaded.resources,
        sheetOrder: reloaded.sheetOrder,
        sheetNames: reloaded.sheetOrder.map((id: string) => reloaded.sheets[id].name),
      };
    });

    expect(result.resources).toEqual([
      { name: 'SHEET_TABLE_PLUGIN', data: JSON.stringify({ tables: [{ id: 't1' }] }) },
      { name: 'SHEET_DATA_VALIDATION_PLUGIN', data: '{}' },
    ]);
    // The hidden meta sheet must NOT appear in the user-visible sheet order.
    expect(result.sheetNames).toEqual(['Sheet1']);
    expect(result.sheetNames).not.toContain('__casual_sheets_resources__');
  });

  test('File → Open replays hyperlinks into the live cell rich-text body', async ({ page }) => {
    const errors: string[] = [];
    page.on('pageerror', (e) => errors.push(`pageerror: ${e.message}`));
    page.on('console', (m) => {
      if (m.type() === 'error') errors.push(`[error] ${m.text()}`);
    });

    const fixture = '/tmp/casual-sheets-hyperlink-fixture.xlsx';
    await writeHyperlinkFixture(page, fixture);

    await page.getByTestId('menubar-file').click();
    const [chooser] = await Promise.all([
      page.waitForEvent('filechooser'),
      page.getByTestId('menu-item-open').click(),
    ]);
    await chooser.setFiles(fixture);

    // The unit swap + AddHyperLinkCommand replay are async. Poll the dev
    // helper (which reads cell.p.body.customRanges off the active snapshot)
    // until both expected links have landed.
    await page.waitForFunction(
      () => (window.__getHyperLinks__?.().length ?? 0) >= 2,
      null,
      { timeout: 10_000 },
    );

    const links = await page.evaluate(() => window.__getHyperLinks__!());
    const payloads = links.map((l) => l.payload).sort();
    expect(payloads).toEqual(
      ['https://anthropic.com', 'https://github.com/schnsrw/sheets'].sort(),
    );
    // Positions survive the round-trip.
    const sortedByCol = [...links].sort((a, b) => a.column - b.column);
    expect(sortedByCol.map((l) => ({ row: l.row, column: l.column }))).toEqual([
      { row: 0, column: 0 },
      { row: 0, column: 1 },
    ]);
    expect(errors, errors.join('\n')).toEqual([]);
  });

  test('File → Save export reads cell rich-text and produces hyperlink cells', async ({ page }) => {
    const fixture = '/tmp/casual-sheets-hyperlink-fixture.xlsx';
    await writeHyperlinkFixture(page, fixture);

    // Open the fixture so HyperLinkModel ends up populated.
    await page.getByTestId('menubar-file').click();
    const [chooser] = await Promise.all([
      page.waitForEvent('filechooser'),
      page.getByTestId('menu-item-open').click(),
    ]);
    await chooser.setFiles(fixture);

    await page.waitForFunction(
      () => (window.__getHyperLinks__?.().length ?? 0) >= 2,
      null,
      { timeout: 10_000 },
    );

    // Trigger the real save path — File → Save dispatches saveAsXlsx because
    // sourceFormat was set to 'xlsx' on open. The save triggers a download
    // which Playwright lets us intercept.
    await page.getByTestId('menubar-file').click();
    const downloadPromise = page.waitForEvent('download');
    await page.getByTestId('menu-item-save').click();
    const download = await downloadPromise;
    const downloadPath = await download.path();
    expect(downloadPath, 'expected a downloaded file path').toBeTruthy();

    const fs = await import('node:fs');
    const buf = fs.readFileSync(downloadPath!);

    // Re-import the saved bytes and confirm the hyperlinks survived the
    // collectExportExtras → workbook write → re-parse loop.
    const pending = await page.evaluate(async (bytes) => {
      const ab = new Uint8Array(bytes).buffer;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const reloaded: any = await window.__xlsx!.xlsxToWorkbookData(ab);
      return reloaded.__pendingHyperlinks ?? [];
    }, Array.from(buf));

    expect(pending).toHaveLength(2);
    const payloads = (pending as Array<{ payload: string }>)
      .map((p) => p.payload)
      .sort();
    expect(payloads).toEqual(
      ['https://anthropic.com', 'https://github.com/schnsrw/sheets'].sort(),
    );
  });
});
