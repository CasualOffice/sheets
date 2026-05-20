import { expect, test, type Page } from '@playwright/test';
import { createRequire } from 'node:module';
import * as path from 'node:path';
import { waitForUniver } from './_helpers';

const require = createRequire(import.meta.url);
// eslint-disable-next-line @typescript-eslint/no-var-requires
const XLSX = require(require.resolve('@e965/xlsx', { paths: [path.join(process.cwd(), 'apps/web')] }));

declare global {
  interface Window {
    __odsNames?: {
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
    window.__odsNames = mod;
  });
}

test.describe('ods defined names round-trip', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await exposeConverters(page);
  });

  test('ods authored with named ranges keeps them in snapshot resources', async ({ page }) => {
    const wb = XLSX.utils.book_new();
    wb.Workbook = { Names: [{ Name: 'Inputs', Ref: 'Data!$A$1:$A$3' }] };
    XLSX.utils.book_append_sheet(
      wb,
      {
        A1: { t: 'n', v: 10 },
        A2: { t: 'n', v: 20 },
        A3: { t: 'n', v: 30 },
        '!ref': 'A1:A3',
      },
      'Data',
    );
    const refBytes = Buffer.from(XLSX.write(wb, { type: 'buffer', bookType: 'ods' }));

    const resources = await page.evaluate(async (bytes) => {
      const buf = new Uint8Array(bytes).buffer;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const snap: any = await window.__odsNames!.odsToWorkbookData(buf);
      return snap.resources as Array<{ name: string; data: string }> | undefined;
    }, Array.from(refBytes));

    const dn = resources?.find((r) => r.name === 'SHEET_DEFINED_NAME_PLUGIN');
    expect(dn, 'SHEET_DEFINED_NAME_PLUGIN resource present').toBeDefined();
    const map = JSON.parse(dn!.data) as Record<string, { name: string; formulaOrRefString: string }>;
    const entry = Object.values(map).find((e) => e.name === 'Inputs');
    expect(entry, 'Inputs named range parsed').toBeDefined();
    expect(entry!.formulaOrRefString).toContain('Data!');
    expect(entry!.formulaOrRefString).toContain('A');
  });

  test('named range we export is visible in ods workbook names', async ({ page }) => {
    const bytes = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const snap: any = {
        id: 'wb-ods-dn-1',
        rev: 1,
        name: 'wb',
        appVersion: '0.22.1',
        locale: 1,
        styles: {},
        sheetOrder: ['s1'],
        sheets: {
          s1: { id: 's1', name: 'Data', cellData: {}, rowCount: 100, columnCount: 26 },
        },
        resources: [
          {
            name: 'SHEET_DEFINED_NAME_PLUGIN',
            data: JSON.stringify({
              'dn-0': {
                id: 'dn-0',
                name: 'Inputs',
                formulaOrRefString: 'Data!$A$1:$A$3',
              },
            }),
          },
        ],
      };
      const blob = await window.__odsNames!.workbookDataToOds(snap);
      return Array.from(new Uint8Array(await blob.arrayBuffer()));
    });

    const wb = XLSX.read(Buffer.from(bytes), { type: 'buffer' });
    const names = (wb.Workbook?.Names ?? []) as Array<{ Name?: string; Ref?: string }>;
    const entry = names.find((n) => n.Name === 'Inputs');
    expect(entry, 'Inputs defined name present in ods').toBeDefined();
    expect(entry!.Ref).toContain('Data!');
    expect(entry!.Ref).toContain('A');
  });
});
