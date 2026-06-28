/**
 * Copyright 2026 Casual Office
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

import { strict as assert } from 'node:assert';
import { test } from 'node:test';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';

import {
  applyDxfCfRulesToZip,
  captureDxfCfRulesFromBuffer,
  captureRawCfFromBuffer,
} from './cf-dxf-passthrough.js';
import {
  dxfCfRulesToSynthCf,
  readDxfCfRulesFromSnapshot,
  applyConditionalFormattingToXlsxWorksheet,
  CONDITIONAL_FORMATTING_RESOURCE,
} from './conditional-formatting-resource.js';

/** A workbook with `numeric A1:A5` on sheet "S"; optionally a cellIs rule so
 *  ExcelJS writes an initial dxf (to exercise index coordination). */
async function baseBuffer(withCellIs = false): Promise<ArrayBuffer> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('S');
  for (let r = 1; r <= 5; r++) ws.getCell(`A${r}`).value = r;
  if (withCellIs) {
    ws.addConditionalFormatting({
      ref: 'A1:A5',
      rules: [
        {
          type: 'cellIs',
          operator: 'greaterThan',
          formulae: ['2'],
          priority: 1,
          style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFF0000' } } },
        },
      ],
    });
  }
  return (await wb.xlsx.writeBuffer()) as ArrayBuffer;
}

const GREEN = { bg: { rgb: '#00ff00' } };

test('export injects a dxf + cfRule, and capture round-trips duplicate/unique', async () => {
  const base = await baseBuffer();
  const zip = await JSZip.loadAsync(base);
  await applyDxfCfRulesToZip(zip, {
    S: [
      { type: 'duplicateValues', sqref: 'A1:A5', style: GREEN },
      { type: 'uniqueValues', sqref: 'B1:B5', style: { bl: 1 } },
    ],
  });
  const out = await zip.generateAsync({ type: 'arraybuffer' });

  const outZip = await JSZip.loadAsync(out);
  const styles = await outZip.file('xl/styles.xml')!.async('string');
  assert.match(styles, /<dxfs count="2">/);
  assert.match(styles, /<bgColor rgb="FF00FF00"\/>/);
  const sheet = await outZip.file('xl/worksheets/sheet1.xml')!.async('string');
  assert.match(sheet, /<cfRule type="duplicateValues" dxfId="0"/);
  assert.match(sheet, /<cfRule type="uniqueValues" dxfId="1"/);

  const captured = await captureDxfCfRulesFromBuffer(out);
  assert.equal(captured?.S?.length, 2);
  const dup = captured!.S.find((r) => r.type === 'duplicateValues')!;
  assert.deepEqual(dup.style, { bg: { rgb: '#00ff00' } });
  const uniq = captured!.S.find((r) => r.type === 'uniqueValues')!;
  assert.deepEqual(uniq.style, { bl: 1 });

  // The app re-imports via ExcelJS, which reconciles each cfRule's dxfId
  // against the styles model — a malformed/duplicate <dxfs> would crash it.
  // Populating the empty <dxfs count="0"/> in place (not adding a second
  // element) keeps the file loadable.
  await assert.doesNotReject(() => new ExcelJS.Workbook().xlsx.load(out as ArrayBuffer));
});

test('dxf index coordinates with dxfs ExcelJS already wrote', async () => {
  // Base already has one dxf (index 0) from the cellIs rule → ours start at 1.
  const base = await baseBuffer(true);
  const zip = await JSZip.loadAsync(base);
  await applyDxfCfRulesToZip(zip, {
    S: [{ type: 'duplicateValues', sqref: 'A1:A5', style: GREEN }],
  });
  const out = await zip.generateAsync({ type: 'arraybuffer' });
  const outZip = await JSZip.loadAsync(out);
  const styles = await outZip.file('xl/styles.xml')!.async('string');
  assert.match(styles, /<dxfs count="2">/); // 1 existing + 1 new
  const sheet = await outZip.file('xl/worksheets/sheet1.xml')!.async('string');
  assert.match(sheet, /<cfRule type="duplicateValues" dxfId="1"/); // appended index
  // Existing cellIs dxf (index 0) is untouched.
  assert.match(sheet, /<cfRule type="cellIs" dxfId="0"/);

  // Capture resolves the new rule's style against dxf index 1.
  const captured = await captureDxfCfRulesFromBuffer(out);
  assert.deepEqual(captured?.S?.[0]?.style, { bg: { rgb: '#00ff00' } });
});

test('snapshot read + synth conversion round-trips duplicate/unique', () => {
  const snapshot = {
    resources: [
      {
        name: CONDITIONAL_FORMATTING_RESOURCE,
        data: JSON.stringify({
          s1: [
            {
              cfId: 'x',
              ranges: [{ startRow: 0, endRow: 4, startColumn: 0, endColumn: 0 }],
              stopIfTrue: false,
              rule: { type: 'highlightCell', subType: 'duplicateValues', style: GREEN },
            },
          ],
        }),
      },
    ],
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
  } as any;
  const bySheetId = readDxfCfRulesFromSnapshot(snapshot);
  assert.deepEqual(bySheetId.s1, [{ type: 'duplicateValues', sqref: 'A1:A5', style: GREEN }]);

  const synth = dxfCfRulesToSynthCf(bySheetId);
  assert.equal(synth.s1?.[0]?.rule.type, 'highlightCell');
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  assert.equal((synth.s1![0].rule as any).subType, 'duplicateValues');
  assert.deepEqual(synth.s1![0].ranges, [{ startRow: 0, endRow: 4, startColumn: 0, endColumn: 0 }]);
});

test('the ExcelJS export path never writes a duplicate/unique rule', async () => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('S');
  for (let r = 1; r <= 5; r++) ws.getCell(`A${r}`).value = r;
  applyConditionalFormattingToXlsxWorksheet(ws, [
    {
      cfId: 'd',
      ranges: [{ startRow: 0, endRow: 4, startColumn: 0, endColumn: 0 }],
      stopIfTrue: false,
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      rule: { type: 'highlightCell', subType: 'duplicateValues', style: GREEN } as any,
    },
  ]);
  const buf = (await wb.xlsx.writeBuffer()) as ArrayBuffer;
  const zip = await JSZip.loadAsync(buf);
  const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
  assert.doesNotMatch(sheet, /duplicateValues/);
});

test('captureRawCfFromBuffer recovers data-bar colours + dup/unique in one pass', async () => {
  // Build a sheet with both a data bar and a duplicateValues rule (ExcelJS
  // writes neither faithfully — inject both via raw XML), then capture once.
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('S');
  for (let r = 1; r <= 5; r++) ws.getCell(`A${r}`).value = r;
  const raw = (await wb.xlsx.writeBuffer()) as ArrayBuffer;
  const zip = await JSZip.loadAsync(raw);
  let styles = await zip.file('xl/styles.xml')!.async('string');
  styles = styles.replace(
    /<dxfs count="0"\/>/,
    '<dxfs count="1"><dxf><fill><patternFill patternType="solid"><bgColor rgb="FF00FF00"/></patternFill></fill></dxf></dxfs>',
  );
  zip.file('xl/styles.xml', styles);
  let sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
  sheet = sheet.replace(
    '<pageMargins',
    '<conditionalFormatting sqref="B1:B5"><cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar></cfRule></conditionalFormatting>' +
      '<conditionalFormatting sqref="A1:A5"><cfRule type="duplicateValues" dxfId="0" priority="2"/></conditionalFormatting>' +
      '<pageMargins',
  );
  zip.file('xl/worksheets/sheet1.xml', sheet);
  const buf = await zip.generateAsync({ type: 'arraybuffer' });

  const { dataBarColors, dxfCfRules } = await captureRawCfFromBuffer(buf);
  assert.equal(dataBarColors?.S?.[0]?.positiveColor, '#638ec6');
  assert.equal(dataBarColors?.S?.[0]?.sqref, 'B1:B5');
  assert.equal(dxfCfRules?.S?.[0]?.type, 'duplicateValues');
  assert.deepEqual(dxfCfRules?.S?.[0]?.style, { bg: { rgb: '#00ff00' } });
});

test('captureRawCfFromBuffer returns empties for a CF-free workbook', async () => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('S');
  for (let r = 1; r <= 5; r++) ws.getCell(`A${r}`).value = r;
  const buf = (await wb.xlsx.writeBuffer()) as ArrayBuffer;
  const out = await captureRawCfFromBuffer(buf);
  assert.equal(out.dataBarColors, undefined);
  assert.equal(out.dxfCfRules, undefined);
});
