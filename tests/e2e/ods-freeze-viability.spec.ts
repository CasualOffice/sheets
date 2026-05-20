import { expect, test } from '@playwright/test';
import { createRequire } from 'node:module';
import * as path from 'node:path';

const require = createRequire(import.meta.url);
// eslint-disable-next-line @typescript-eslint/no-var-requires
const XLSX = require(require.resolve('@e965/xlsx', { paths: [path.join(process.cwd(), 'apps/web')] }));

test.describe('ods freeze-pane viability probes', () => {
  test('current SheetJS ODS writer does not emit freeze metadata into content.xml', async () => {
    const ws = { A1: { t: 's', v: 'x' }, '!ref': 'A1' };
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Data');
    wb.Workbook = { Views: [{ xSplit: 1, ySplit: 1 }] };

    const out = XLSX.write(wb, { type: 'buffer', bookType: 'ods' });
    const cfb = XLSX.CFB.read(out, { type: 'buffer' });
    const entry = cfb.FileIndex.find((f: { name?: string }) => f.name === 'content.xml');
    expect(entry, 'content.xml entry present').toBeDefined();
    const contentXml = Buffer.from(entry.content).toString('utf8');

    expect(contentXml.toLowerCase()).not.toContain('freezepanes');
    expect(contentXml.toLowerCase()).not.toContain('splithorizontal');
    expect(contentXml.toLowerCase()).not.toContain('splitvertical');
  });

  test('current SheetJS ODS parser does not surface workbook or sheet freeze views', async () => {
    const ws = { A1: { t: 's', v: 'x' }, '!ref': 'A1' };
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Data');
    wb.Workbook = { Views: [{ xSplit: 1, ySplit: 1 }] };

    const out = XLSX.write(wb, { type: 'buffer', bookType: 'ods' });
    const back = XLSX.read(out, { type: 'buffer', cellFormula: true, cellNF: true });

    expect(back.Workbook?.Views).toBeUndefined();
    expect(back.Sheets.Data['!freeze']).toBeUndefined();
  });
});
