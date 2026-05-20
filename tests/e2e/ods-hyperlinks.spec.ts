import { expect, test, type Page } from '@playwright/test';
import { createRequire } from 'node:module';
import * as path from 'node:path';
import { waitForUniver } from './_helpers';

const require = createRequire(import.meta.url);
// eslint-disable-next-line @typescript-eslint/no-var-requires
const XLSX = require(require.resolve('@e965/xlsx', { paths: [path.join(process.cwd(), 'apps/web')] }));

declare global {
  interface Window {
    __odsLinks?: {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      odsToWorkbookData: (buf: ArrayBuffer) => Promise<any>;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      workbookDataToOds: (data: any) => Promise<Blob>;
    };
  }
}

async function exposeConverters(page: Page) {
  await page.evaluate(async () => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const mod = await import(/* @vite-ignore */ '/src/ods/index.ts' as any);
    window.__odsLinks = mod;
  });
}

test.describe('ods hyperlinks round-trip', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await exposeConverters(page);
  });

  test('importer encodes cell hyperlinks inline in cell.p.body', async ({ page }) => {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(
      wb,
      {
        A1: { t: 's', v: 'GitHub', l: { Target: 'https://github.com/schnsrw/sheets' } },
        '!ref': 'A1',
      },
      'Links',
    );
    const refBytes = Buffer.from(XLSX.write(wb, { type: 'buffer', bookType: 'ods' }));

    const probe = await page.evaluate(async (bytes) => {
      const buf = new Uint8Array(bytes).buffer;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const snap: any = await window.__odsLinks!.odsToWorkbookData(buf);
      const firstSheetId = snap.sheetOrder[0];
      const cell = snap.sheets[firstSheetId].cellData[0][0];
      return {
        v: cell.v,
        body: cell.p?.body ?? null,
      };
    }, Array.from(refBytes));

    expect(probe.v).toBe('GitHub');
    expect(probe.body).not.toBeNull();
    expect(probe.body.dataStream.startsWith('GitHub')).toBe(true);
    const cr = probe.body.customRanges?.[0];
    expect(cr).toBeTruthy();
    expect(cr.rangeType).toBe(0);
    expect(cr.properties?.url).toBe('https://github.com/schnsrw/sheets');
  });

  test('hyperlink in snapshot exports as ods cell hyperlink', async ({ page }) => {
    const bytes = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const snap: any = {
        id: 'wb-ods-link-1',
        rev: 1,
        name: 'wb',
        appVersion: '0.22.1',
        locale: 1,
        styles: {},
        sheetOrder: ['s1'],
        sheets: {
          s1: {
            id: 's1',
            name: 'Links',
            cellData: {
              0: {
                0: {
                  v: 'GitHub',
                  p: {
                    id: '__INTERNAL_EDITOR__DOCS_NORMAL',
                    documentStyle: {},
                    body: {
                      dataStream: 'GitHub\r\n',
                      customRanges: [
                        {
                          startIndex: 0,
                          endIndex: 5,
                          rangeType: 0,
                          rangeId: 'hl-1',
                          properties: { url: 'https://github.com/schnsrw/sheets' },
                        },
                      ],
                      paragraphs: [{ startIndex: 6 }],
                      sectionBreaks: [{ startIndex: 7 }],
                      textRuns: [],
                    },
                  },
                },
              },
            },
            rowCount: 100,
            columnCount: 26,
          },
        },
      };
      const blob = await window.__odsLinks!.workbookDataToOds(snap);
      return Array.from(new Uint8Array(await blob.arrayBuffer()));
    });

    const wb = XLSX.read(Buffer.from(bytes), { type: 'buffer' });
    expect(wb.Sheets.Links.A1?.l?.Target).toBe('https://github.com/schnsrw/sheets');
  });
});
