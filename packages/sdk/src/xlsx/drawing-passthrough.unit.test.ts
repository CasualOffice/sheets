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

import { applyDrawingsToZip, captureDrawingsFromBuffer } from './drawing-passthrough.js';

// A 1×1 transparent PNG — enough for ExcelJS to emit xl/media + xl/drawings.
const PNG_1X1_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

/** Build an xlsx buffer with a single PNG image anchored on a sheet. */
async function bufferWithImage(sheetName = 'S'): Promise<ArrayBuffer> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(sheetName);
  ws.getCell('A1').value = 'has image';
  const imageId = wb.addImage({ base64: PNG_1X1_BASE64, extension: 'png' });
  ws.addImage(imageId, 'B2:D6');
  return (await wb.xlsx.writeBuffer()) as ArrayBuffer;
}

test('captureDrawingsFromBuffer recovers media + drawing parts and the sheet linkage', async () => {
  const buf = await bufferWithImage('S');
  const payload = await captureDrawingsFromBuffer(buf);

  assert.ok(payload, 'expected a drawing payload');
  const paths = Object.keys(payload!.parts);
  assert.ok(
    paths.some((p) => /^xl\/media\/image\d+\.png$/.test(p)),
    `expected a media png part, got ${paths.join(', ')}`,
  );
  assert.ok(
    paths.some((p) => /^xl\/drawings\/drawing\d+\.xml$/.test(p)),
    `expected a drawing xml part, got ${paths.join(', ')}`,
  );
  // The sheet→drawing linkage is keyed by decoded sheet name.
  assert.equal(payload!.perSheet.length, 1);
  assert.equal(payload!.perSheet[0].sheetName, 'S');
  assert.match(payload!.perSheet[0].drawingTarget, /^xl\/drawings\/drawing\d+\.xml$/);
});

test('a workbook with no images yields no payload', async () => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Plain');
  ws.getCell('A1').value = 1;
  const buf = (await wb.xlsx.writeBuffer()) as ArrayBuffer;
  assert.equal(await captureDrawingsFromBuffer(buf), undefined);
});

test('applyDrawingsToZip restores the image so ExcelJS reads it back (full round-trip)', async () => {
  // 1. Capture from a real image workbook.
  const payload = (await captureDrawingsFromBuffer(await bufferWithImage('S')))!;
  assert.ok(payload);

  // 2. A fresh ExcelJS export of the same sheet drops the image entirely
  //    (Univer has no drawing model). Confirm the baseline really is empty.
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('S');
  ws.getCell('A1').value = 'has image';
  const baseBuf = (await wb.xlsx.writeBuffer()) as ArrayBuffer;
  {
    const check = new ExcelJS.Workbook();
    await check.xlsx.load(baseBuf);
    assert.equal(check.getWorksheet('S')!.getImages().length, 0, 'baseline should have no image');
  }

  // 3. Splice the captured drawing parts back in.
  const zip = await JSZip.loadAsync(baseBuf);
  await applyDrawingsToZip(zip, payload);
  const out = await zip.generateAsync({ type: 'arraybuffer' });

  // 4. ExcelJS must now read the image back, proving the parts + content-types
  //    + sheet <drawing> rel were all re-linked correctly.
  const reload = new ExcelJS.Workbook();
  await reload.xlsx.load(out);
  assert.equal(reload.getWorksheet('S')!.getImages().length, 1, 'image should survive round-trip');

  // Content types declare the png default + the drawing override.
  const ct = await (await JSZip.loadAsync(out)).file('[Content_Types].xml')!.async('string');
  assert.match(ct, /<Default Extension="png"/);
  assert.match(ct, /drawing\+xml/);
});

test('apply matches the sheet by decoded name (XML-special chars)', async () => {
  const SHEET = 'R&D <Q1>';
  const payload = (await captureDrawingsFromBuffer(await bufferWithImage(SHEET)))!;
  assert.ok(payload);
  assert.equal(payload.perSheet[0].sheetName, SHEET);

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(SHEET);
  ws.getCell('A1').value = 'x';
  const baseBuf = (await wb.xlsx.writeBuffer()) as ArrayBuffer;
  const zip = await JSZip.loadAsync(baseBuf);
  await applyDrawingsToZip(zip, payload);
  const out = await zip.generateAsync({ type: 'arraybuffer' });

  const reload = new ExcelJS.Workbook();
  await reload.xlsx.load(out);
  assert.equal(reload.getWorksheet(SHEET)!.getImages().length, 1);
});
