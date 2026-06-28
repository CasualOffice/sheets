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
  applyThreadedCommentsToZip,
  captureThreadedCommentsFromBuffer,
} from './threaded-comment-passthrough.js';

const PERSON_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments" ' +
  'xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
  '<person displayName="Ada Lovelace" id="{AAAA0000-0000-0000-0000-000000000001}" userId="ada" providerId="None"/>' +
  '<person displayName="Bob Stone" id="{BBBB0000-0000-0000-0000-000000000002}" userId="bob" providerId="None"/>' +
  '</personList>';

const THREAD_XML = (sheetSuffix: string) =>
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">' +
  `<threadedComment ref="B2" dT="2026-01-01T10:00:00Z" personId="{AAAA0000-0000-0000-0000-000000000001}" id="{C${sheetSuffix}00-0000-0000-0000-000000000001}"><text>Looks off here</text></threadedComment>` +
  `<threadedComment ref="B2" dT="2026-01-01T11:00:00Z" personId="{BBBB0000-0000-0000-0000-000000000002}" id="{C${sheetSuffix}00-0000-0000-0000-000000000002}" parentId="{C${sheetSuffix}00-0000-0000-0000-000000000001}"><text>Agreed, fixing</text></threadedComment>` +
  '</ThreadedComments>';

/** Hand-build an xlsx with a 2-message threaded comment + persons on one sheet. */
async function bufferWithThread(sheetName = 'S'): Promise<ArrayBuffer> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(sheetName);
  // A legacy note (ExcelJS models this) backs the threaded comment.
  ws.getCell('B2').note = 'Looks off here';
  const base = (await wb.xlsx.writeBuffer()) as ArrayBuffer;

  const zip = await JSZip.loadAsync(base);
  zip.file('xl/persons/person.xml', PERSON_XML);
  zip.file('xl/threadedComments/threadedComment1.xml', THREAD_XML('1'));

  let ct = await zip.file('[Content_Types].xml')!.async('string');
  ct = ct.replace(
    '</Types>',
    '<Override PartName="/xl/persons/person.xml" ContentType="application/vnd.ms-excel.person+xml"/>' +
      '<Override PartName="/xl/threadedComments/threadedComment1.xml" ContentType="application/vnd.ms-excel.threadedcomments+xml"/></Types>',
  );
  zip.file('[Content_Types].xml', ct);

  let wbRels = await zip.file('xl/_rels/workbook.xml.rels')!.async('string');
  wbRels = wbRels.replace(
    '</Relationships>',
    '<Relationship Id="rId800" Type="http://schemas.microsoft.com/office/2017/10/relationships/person" Target="persons/person.xml"/></Relationships>',
  );
  zip.file('xl/_rels/workbook.xml.rels', wbRels);

  // The sheet rels ExcelJS wrote for the legacy comment — add the threadedComment rel.
  const sheetRelsPath = 'xl/worksheets/_rels/sheet1.xml.rels';
  let sheetRels =
    (await zip.file(sheetRelsPath)?.async('string')) ??
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  sheetRels = sheetRels.replace(
    '</Relationships>',
    '<Relationship Id="rId900" Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment" Target="../threadedComments/threadedComment1.xml"/></Relationships>',
  );
  zip.file(sheetRelsPath, sheetRels);

  return zip.generateAsync({ type: 'arraybuffer' });
}

test('capture recovers threaded-comment + persons parts and the sheet linkage', async () => {
  const payload = await captureThreadedCommentsFromBuffer(await bufferWithThread('S'));
  assert.ok(payload, 'expected a threaded-comment payload');
  assert.ok(payload!.parts['xl/threadedComments/threadedComment1.xml']);
  assert.ok(payload!.parts['xl/persons/person.xml']);
  assert.equal(payload!.perSheet.length, 1);
  assert.equal(payload!.perSheet[0].sheetName, 'S');
  assert.equal(payload!.perSheet[0].target, 'xl/threadedComments/threadedComment1.xml');
});

test('a workbook with only a legacy note (no thread) yields no payload', async () => {
  const wb = new ExcelJS.Workbook();
  wb.addWorksheet('S').getCell('A1').note = 'plain note';
  const buf = (await wb.xlsx.writeBuffer()) as ArrayBuffer;
  assert.equal(await captureThreadedCommentsFromBuffer(buf), undefined);
});

test('apply restores parts + content-types + relationships, and ExcelJS still loads', async () => {
  const payload = (await captureThreadedCommentsFromBuffer(await bufferWithThread('S')))!;
  assert.ok(payload);

  // A fresh ExcelJS export keeps the legacy note but drops the threaded layer.
  const wb = new ExcelJS.Workbook();
  wb.addWorksheet('S').getCell('B2').note = 'Looks off here';
  const base = (await wb.xlsx.writeBuffer()) as ArrayBuffer;
  assert.equal(
    (await JSZip.loadAsync(base)).file('xl/threadedComments/threadedComment1.xml'),
    null,
    'baseline should have no threaded comment',
  );

  const zip = await JSZip.loadAsync(base);
  await applyThreadedCommentsToZip(zip, payload);
  const out = await zip.generateAsync({ type: 'arraybuffer' });
  const outZip = await JSZip.loadAsync(out);

  // Parts restored, with the real authors + reply chain intact.
  const thread = await outZip.file('xl/threadedComments/threadedComment1.xml')!.async('string');
  assert.match(thread, /parentId=/, 'reply (parentId) preserved');
  const persons = await outZip.file('xl/persons/person.xml')!.async('string');
  assert.match(persons, /Ada Lovelace/);
  assert.match(persons, /Bob Stone/);

  // Content types declared.
  const ct = await outZip.file('[Content_Types].xml')!.async('string');
  assert.match(ct, /person\+xml/);
  assert.match(ct, /threadedcomments\+xml/);

  // Relationships re-created: workbook→persons and sheet→threadedComment.
  const wbRels = await outZip.file('xl/_rels/workbook.xml.rels')!.async('string');
  assert.match(wbRels, /relationships\/person"[^>]*Target="persons\/person\.xml"/);
  const sheetRels = await outZip.file('xl/worksheets/_rels/sheet1.xml.rels')!.async('string');
  assert.match(
    sheetRels,
    /relationships\/threadedComment"[^>]*Target="\.\.\/threadedComments\/threadedComment1\.xml"/,
  );

  // ExcelJS still loads the result (it ignores threaded parts, must not throw).
  const reload = new ExcelJS.Workbook();
  await reload.xlsx.load(out);
  assert.ok(reload.getWorksheet('S'));

  // Re-capture round-trips.
  const again = await captureThreadedCommentsFromBuffer(out);
  assert.equal(again?.perSheet[0]?.sheetName, 'S');
});

test('apply matches the sheet by decoded name (XML-special chars)', async () => {
  const SHEET = 'R&D <Q1>';
  const payload = (await captureThreadedCommentsFromBuffer(await bufferWithThread(SHEET)))!;
  assert.ok(payload);
  assert.equal(payload.perSheet[0].sheetName, SHEET);

  const wb = new ExcelJS.Workbook();
  wb.addWorksheet(SHEET).getCell('B2').note = 'Looks off here';
  const base = (await wb.xlsx.writeBuffer()) as ArrayBuffer;
  const zip = await JSZip.loadAsync(base);
  await applyThreadedCommentsToZip(zip, payload);
  const out = await zip.generateAsync({ type: 'arraybuffer' });

  const outZip = await JSZip.loadAsync(out);
  const sheetRels = await outZip.file('xl/worksheets/_rels/sheet1.xml.rels')!.async('string');
  assert.match(sheetRels, /relationships\/threadedComment"/);
});
