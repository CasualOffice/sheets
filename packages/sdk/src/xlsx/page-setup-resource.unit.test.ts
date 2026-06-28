import { strict as assert } from 'node:assert';
import { test } from 'node:test';
import ExcelJS from 'exceljs';

import {
  applyPageSetupToXlsxWorksheet,
  readPageSetupFromXlsx,
  type SheetPageSetupV1,
} from './page-setup-resource.js';

const idFor = (excelId: number): string => `sheet-${excelId}`;

test('readPageSetupFromXlsx captures print titles (repeat rows/cols)', async () => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('S');
  ws.getCell('A1').value = 'h';
  ws.pageSetup.printTitlesRow = '1:1'; // repeat row 1 at the top of every page
  ws.pageSetup.printTitlesColumn = 'A:A'; // repeat column A on the left
  ws.pageSetup.printArea = 'A1:C10';

  const reloaded = new ExcelJS.Workbook();
  await reloaded.xlsx.load((await wb.xlsx.writeBuffer()) as ArrayBuffer);
  const map = readPageSetupFromXlsx(reloaded, idFor);
  const entry = Object.values(map)[0];
  assert.ok(entry, 'expected a page-setup entry');
  assert.equal(entry.printTitlesRow, '1:1');
  assert.equal(entry.printTitlesColumn, 'A:A');
  assert.equal(entry.printArea, 'A1:C10');
});

test('applyPageSetupToXlsxWorksheet writes print titles onto a worksheet', () => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('S');
  const entry: SheetPageSetupV1 = { printTitlesRow: '1:2', printTitlesColumn: 'A:B' };
  applyPageSetupToXlsxWorksheet(ws, entry);
  assert.equal(ws.pageSetup.printTitlesRow, '1:2');
  assert.equal(ws.pageSetup.printTitlesColumn, 'A:B');
});

test('print titles survive a full read → apply → re-read round-trip', async () => {
  // Source file with print titles.
  const src = new ExcelJS.Workbook();
  const sws = src.addWorksheet('S');
  sws.pageSetup.printTitlesRow = '1:1';
  sws.pageSetup.printTitlesColumn = 'A:A';
  const loaded = new ExcelJS.Workbook();
  await loaded.xlsx.load((await src.xlsx.writeBuffer()) as ArrayBuffer);
  const entry = Object.values(readPageSetupFromXlsx(loaded, idFor))[0];

  // Export side: apply the captured entry onto a freshly-built workbook.
  const out = new ExcelJS.Workbook();
  applyPageSetupToXlsxWorksheet(out.addWorksheet('S'), entry);
  const roundTripped = new ExcelJS.Workbook();
  await roundTripped.xlsx.load((await out.xlsx.writeBuffer()) as ArrayBuffer);
  const after = Object.values(readPageSetupFromXlsx(roundTripped, idFor))[0];
  assert.equal(after.printTitlesRow, '1:1');
  assert.equal(after.printTitlesColumn, 'A:A');
});

test('no print titles → no fields added (no spurious sidecar)', async () => {
  const wb = new ExcelJS.Workbook();
  wb.addWorksheet('Plain').getCell('A1').value = 1;
  const reloaded = new ExcelJS.Workbook();
  await reloaded.xlsx.load((await wb.xlsx.writeBuffer()) as ArrayBuffer);
  const entry = Object.values(readPageSetupFromXlsx(reloaded, idFor))[0];
  // entry may be undefined (nothing non-default) or present without titles.
  assert.equal(entry?.printTitlesRow, undefined);
  assert.equal(entry?.printTitlesColumn, undefined);
});
