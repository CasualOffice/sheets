import { strict as assert } from 'node:assert';
import { test } from 'node:test';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';

import { applyDataBarsToZip, captureDataBarColorsFromBuffer } from './databar-passthrough.js';
import {
  applyConditionalFormattingToXlsxWorksheet,
  readConditionalFormattingFromXlsx,
  readDataBarsFromSnapshot,
  CONDITIONAL_FORMATTING_RESOURCE,
} from './conditional-formatting-resource.js';

/** Build an xlsx buffer with one data bar over A1:A5 on sheet "S". */
async function bufferWithDataBar(argb = 'FF638EC6'): Promise<ArrayBuffer> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('S');
  for (let r = 1; r <= 5; r++) ws.getCell(`A${r}`).value = r * 10;
  ws.addConditionalFormatting({
    ref: 'A1:A5',
    rules: [
      { type: 'dataBar', cfvo: [{ type: 'min' }, { type: 'max' }], color: [{ argb }], priority: 1 },
    ],
  });
  return (await wb.xlsx.writeBuffer()) as ArrayBuffer;
}

test('captureDataBarColorsFromBuffer recovers the positive fill ExcelJS drops', async () => {
  // ExcelJS itself writes <color auto="1"/>, so capture from an ExcelJS buffer
  // returns nothing — exercise the parser against a hand-written rgb color.
  const buf = await bufferWithDataBar();
  const zip = await JSZip.loadAsync(buf);
  let xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
  xml = xml.replace('<color auto="1"/>', '<color rgb="FF638EC6"/>');
  zip.file('xl/worksheets/sheet1.xml', xml);
  const patched = await zip.generateAsync({ type: 'arraybuffer' });

  const colors = await captureDataBarColorsFromBuffer(patched);
  assert.deepEqual(colors, { S: [{ sqref: 'A1:A5', positiveColor: '#638ec6' }] });
});

test('import maps a data bar (ExcelJS shape + injected color) to an IDataBar rule', async () => {
  const buf = await bufferWithDataBar();
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);
  const cf = readConditionalFormattingFromXlsx(wb, (id) => `sheet-${id}`, {
    S: [{ sqref: 'A1:A5', positiveColor: '#638ec6' }],
  });
  const rules = Object.values(cf)[0] ?? [];
  assert.equal(rules.length, 1);
  assert.deepEqual(rules[0].rule, {
    type: 'dataBar',
    isShowValue: true,
    config: {
      min: { type: 'min' },
      max: { type: 'max' },
      isGradient: true,
      positiveColor: '#638ec6',
      nativeColor: '#ff0000',
    },
  });
});

test('a data bar with no recovered colour falls back to Excel default blue', async () => {
  const buf = await bufferWithDataBar();
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);
  // No colour map passed → default.
  const cf = readConditionalFormattingFromXlsx(wb, (id) => `sheet-${id}`);
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const rule = (Object.values(cf)[0]?.[0]?.rule ?? {}) as any;
  assert.equal(rule.config?.positiveColor, '#638ec6');
});

test('export splices a data bar into the worksheet XML with the real colour', async () => {
  // Snapshot carrying one IDataBar in the CF resource.
  const snapshot = {
    sheetOrder: ['s1'],
    sheets: { s1: { name: 'S' } },
    resources: [
      {
        name: CONDITIONAL_FORMATTING_RESOURCE,
        data: JSON.stringify({
          s1: [
            {
              cfId: 'd',
              ranges: [{ startRow: 0, endRow: 4, startColumn: 0, endColumn: 0 }],
              stopIfTrue: false,
              rule: {
                type: 'dataBar',
                isShowValue: true,
                config: {
                  min: { type: 'min' },
                  max: { type: 'max' },
                  isGradient: true,
                  positiveColor: '#12ab34',
                  nativeColor: '#ff0000',
                },
              },
            },
          ],
        }),
      },
    ],
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
  } as any;

  const entries = readDataBarsFromSnapshot(snapshot);
  assert.deepEqual(entries.s1?.[0], {
    sqref: 'A1:A5',
    positiveColor: '#12ab34',
    isShowValue: true,
    min: { type: 'min' },
    max: { type: 'max' },
  });

  // Apply onto a real ExcelJS-written zip for sheet "S".
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('S');
  for (let r = 1; r <= 5; r++) ws.getCell(`A${r}`).value = r * 10;
  const baseBuf = (await wb.xlsx.writeBuffer()) as ArrayBuffer;
  const zip = await JSZip.loadAsync(baseBuf);
  await applyDataBarsToZip(zip, { S: entries.s1 });
  const out = await zip.generateAsync({ type: 'arraybuffer' });

  const outZip = await JSZip.loadAsync(out);
  const xml = await outZip.file('xl/worksheets/sheet1.xml')!.async('string');
  assert.match(xml, /<cfRule type="dataBar"[^>]*>/);
  assert.match(xml, /<color rgb="FF12AB34"\/>/);
  assert.match(xml, /sqref="A1:A5"/);

  // Round-trips: re-capture recovers the colour.
  const colors = await captureDataBarColorsFromBuffer(out);
  assert.equal(colors?.S?.[0]?.positiveColor, '#12ab34');
});

test('sheet names with XML-special chars round-trip (capture + apply key by decoded name)', async () => {
  const SHEET = 'R&D <Q1>';
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(SHEET);
  for (let r = 1; r <= 5; r++) ws.getCell(`A${r}`).value = r * 10;
  ws.addConditionalFormatting({
    ref: 'A1:A5',
    rules: [{ type: 'dataBar', cfvo: [{ type: 'min' }, { type: 'max' }], priority: 1 }],
  });
  const raw = (await wb.xlsx.writeBuffer()) as ArrayBuffer;
  const zip = await JSZip.loadAsync(raw);
  let xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
  // Give the data bar a real positive colour inside <dataBar> (ExcelJS omits one).
  xml = xml.replace('</dataBar>', '<color rgb="FF112233"/></dataBar>');
  zip.file('xl/worksheets/sheet1.xml', xml);
  const buf = await zip.generateAsync({ type: 'arraybuffer' });

  // Capture keys by the DECODED sheet name (not "R&amp;D &lt;Q1&gt;").
  const colors = await captureDataBarColorsFromBuffer(buf);
  assert.equal(colors?.[SHEET]?.[0]?.positiveColor, '#112233');

  // Apply finds the sheet by its decoded name too.
  const z2 = await JSZip.loadAsync(buf);
  await applyDataBarsToZip(z2, {
    [SHEET]: [
      {
        sqref: 'A1:A5',
        positiveColor: '#445566',
        isShowValue: true,
        min: { type: 'min' },
        max: { type: 'max' },
      },
    ],
  });
  const out = await z2.generateAsync({ type: 'arraybuffer' });
  const outZip = await JSZip.loadAsync(out);
  const outXml = await outZip.file('xl/worksheets/sheet1.xml')!.async('string');
  assert.match(outXml, /<color rgb="FF445566"\/>/);
});

test('the ExcelJS export path never writes a data bar (only the raw-XML splice does)', async () => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('S');
  for (let r = 1; r <= 5; r++) ws.getCell(`A${r}`).value = r * 10;
  // Feed a dataBar synth rule to the ExcelJS applier — it must add nothing.
  applyConditionalFormattingToXlsxWorksheet(ws, [
    {
      cfId: 'd',
      ranges: [{ startRow: 0, endRow: 4, startColumn: 0, endColumn: 0 }],
      stopIfTrue: false,
      rule: {
        type: 'dataBar',
        isShowValue: true,
        config: {
          min: { type: 'min' },
          max: { type: 'max' },
          isGradient: true,
          positiveColor: '#12ab34',
          nativeColor: '#ff0000',
        },
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
      } as any,
    },
  ]);
  const buf = (await wb.xlsx.writeBuffer()) as ArrayBuffer;
  const zip = await JSZip.loadAsync(buf);
  const xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
  assert.doesNotMatch(xml, /type="dataBar"/);
});
