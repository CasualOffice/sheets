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

import { generateNativePivot } from './pivot-export.js';
import { applyPivotsToZip, capturePivotsFromBuffer } from './pivot-passthrough.js';

const SAMPLE = {
  sheetName: 'Sheet1',
  sourceRef: 'A1:C5',
  headers: ['Region', 'Quarter', 'Sales'],
  records: [
    ['North', 'Q1', 100],
    ['South', 'Q1', 80],
    ['North', 'Q2', 120],
    ['South', 'Q2', 95],
  ] as Array<Array<string | number | null>>,
  rowField: 0,
  valueField: 2,
  valueAgg: 'sum' as const,
  targetRef: 'E1',
};

const decode = (b64: string) => Buffer.from(b64, 'base64').toString('utf8');

test('generates the three pivot parts + their rels', () => {
  const p = generateNativePivot(SAMPLE);
  assert.ok(p.parts['xl/pivotCaches/pivotCacheDefinition1.xml']);
  assert.ok(p.parts['xl/pivotCaches/pivotCacheRecords1.xml']);
  assert.ok(p.parts['xl/pivotTables/pivotTable1.xml']);
  assert.ok(p.parts['xl/pivotCaches/_rels/pivotCacheDefinition1.xml.rels']);
  assert.ok(p.parts['xl/pivotTables/_rels/pivotTable1.xml.rels']);
});

test('cacheDefinition points at the source range + refreshes on load', () => {
  const def = decode(generateNativePivot(SAMPLE).parts['xl/pivotCaches/pivotCacheDefinition1.xml']);
  assert.match(def, /<worksheetSource ref="A1:C5" sheet="Sheet1"\/>/);
  assert.match(def, /refreshOnLoad="1"/);
  // Region (string) gets shared items; Sales (numeric) gets number flags.
  assert.match(
    def,
    /<cacheField name="Region"[^>]*><sharedItems count="2"><s v="North"\/><s v="South"\/>/,
  );
  assert.match(
    def,
    /<cacheField name="Sales"[^>]*><sharedItems[^>]*containsNumber="1"[^>]*minValue="80" maxValue="120"/,
  );
});

test('cacheRecords encode strings by shared index and numbers inline', () => {
  const rec = decode(generateNativePivot(SAMPLE).parts['xl/pivotCaches/pivotCacheRecords1.xml']);
  assert.match(rec, /count="4"/);
  // First record: North (idx 0), Q1 (idx 0), 100 inline.
  assert.match(rec, /<r><x v="0"\/><x v="0"\/><n v="100"\/><\/r>/);
  // South is shared index 1 in Region.
  assert.match(rec, /<r><x v="1"\/>/);
});

test('pivotTable wires the row field, data field, and output location', () => {
  const pt = decode(generateNativePivot(SAMPLE).parts['xl/pivotTables/pivotTable1.xml']);
  assert.match(pt, /<location ref="E1:F4"/); // header + 2 regions + grand total
  assert.match(pt, /<rowFields count="1"><field x="0"\/><\/rowFields>/);
  assert.match(pt, /<dataField name="Sum of Sales" fld="2"/);
  assert.match(pt, /axis="axisRow"/);
  assert.match(pt, /<i t="grand">/);
});

test('non-sum aggregations set the dataField subtotal', () => {
  const pt = decode(
    generateNativePivot({ ...SAMPLE, valueAgg: 'average' }).parts['xl/pivotTables/pivotTable1.xml'],
  );
  assert.match(pt, /<dataField name="Average of Sales"[^>]*subtotal="average"/);
});

test('round-trips: inject into a real xlsx, then re-capture the parts', async () => {
  // Build a minimal base workbook (source + materialised pivot cells).
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Sheet1');
  ws.addRow(['Region', 'Quarter', 'Sales']);
  for (const r of SAMPLE.records) ws.addRow(r as Array<string | number>);
  ws.getCell('E1').value = 'Region';
  ws.getCell('F1').value = 'Sum of Sales';
  const base = (await wb.xlsx.writeBuffer()) as ArrayBuffer;

  const zip = await JSZip.loadAsync(base);
  await applyPivotsToZip(zip, generateNativePivot(SAMPLE));
  const out = await zip.generateAsync({ type: 'arraybuffer' });

  // The injected file must re-capture as a pivot workbook (proves the parts,
  // workbook <pivotCaches>, and sheet rels are all wired correctly).
  const recaptured = await capturePivotsFromBuffer(out);
  assert.ok(recaptured, 'expected pivot parts to be re-captured');
  assert.ok(recaptured!.parts['xl/pivotTables/pivotTable1.xml']);
  assert.equal(recaptured!.perSheet[0]?.sheetName, 'Sheet1');
  assert.ok(recaptured!.workbookCacheRels.length >= 1);
});
