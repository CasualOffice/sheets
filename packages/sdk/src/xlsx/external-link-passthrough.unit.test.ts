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
  applyExternalLinksToZip,
  captureExternalLinksFromBuffer,
} from './external-link-passthrough.js';

const LINK_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
  '<externalBook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1">' +
  '<sheetNames><sheetName val="Sheet1"/></sheetNames>' +
  '<sheetDataSet><sheetData sheetId="0"><row r="1"><cell r="A1"><v>42</v></cell></row></sheetData></sheetDataSet>' +
  '</externalBook></externalLink>';

const LINK_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" ' +
  'Target="file:///C:/Book2.xlsx" TargetMode="External"/></Relationships>';

/** Hand-build an xlsx with one external link (ExcelJS can't author them). */
async function bufferWithExternalLink(): Promise<ArrayBuffer> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('S');
  ws.getCell('A1').value = { formula: '[1]Sheet1!A1', result: 42 };
  const base = (await wb.xlsx.writeBuffer()) as ArrayBuffer;

  const zip = await JSZip.loadAsync(base);
  zip.file('xl/externalLinks/externalLink1.xml', LINK_XML);
  zip.file('xl/externalLinks/_rels/externalLink1.xml.rels', LINK_RELS);

  let ct = await zip.file('[Content_Types].xml')!.async('string');
  ct = ct.replace(
    '</Types>',
    '<Override PartName="/xl/externalLinks/externalLink1.xml" ' +
      'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/></Types>',
  );
  zip.file('[Content_Types].xml', ct);

  let rels = await zip.file('xl/_rels/workbook.xml.rels')!.async('string');
  rels = rels.replace(
    '</Relationships>',
    '<Relationship Id="rId900" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" ' +
      'Target="externalLinks/externalLink1.xml"/></Relationships>',
  );
  zip.file('xl/_rels/workbook.xml.rels', rels);

  let wbXml = await zip.file('xl/workbook.xml')!.async('string');
  const block = '<externalReferences><externalReference r:id="rId900"/></externalReferences>';
  wbXml = wbXml.includes('<definedNames')
    ? wbXml.replace('<definedNames', `${block}<definedNames`)
    : wbXml.replace('</workbook>', `${block}</workbook>`);
  zip.file('xl/workbook.xml', wbXml);

  return zip.generateAsync({ type: 'arraybuffer' });
}

test('captureExternalLinksFromBuffer recovers the link parts in reference order', async () => {
  const payload = await captureExternalLinksFromBuffer(await bufferWithExternalLink());
  assert.ok(payload, 'expected an external-link payload');
  assert.ok(payload!.parts['xl/externalLinks/externalLink1.xml']);
  assert.ok(payload!.parts['xl/externalLinks/_rels/externalLink1.xml.rels']);
  assert.deepEqual(payload!.order, ['xl/externalLinks/externalLink1.xml']);
});

test('a workbook with no external links yields no payload', async () => {
  const wb = new ExcelJS.Workbook();
  wb.addWorksheet('Plain').getCell('A1').value = 1;
  const buf = (await wb.xlsx.writeBuffer()) as ArrayBuffer;
  assert.equal(await captureExternalLinksFromBuffer(buf), undefined);
});

test('apply restores parts + content-types + workbook references (full round-trip)', async () => {
  const payload = (await captureExternalLinksFromBuffer(await bufferWithExternalLink()))!;
  assert.ok(payload);

  // A fresh ExcelJS export drops external links entirely.
  const wb = new ExcelJS.Workbook();
  wb.addWorksheet('S').getCell('A1').value = { formula: '[1]Sheet1!A1', result: 42 };
  const base = (await wb.xlsx.writeBuffer()) as ArrayBuffer;
  assert.equal(
    Object.keys(await JSZip.loadAsync(base)).length >= 0 &&
      (await JSZip.loadAsync(base)).file('xl/externalLinks/externalLink1.xml'),
    null,
    'baseline should have no external link',
  );

  const zip = await JSZip.loadAsync(base);
  await applyExternalLinksToZip(zip, payload);
  const out = await zip.generateAsync({ type: 'arraybuffer' });

  const outZip = await JSZip.loadAsync(out);
  // Part restored.
  assert.ok(outZip.file('xl/externalLinks/externalLink1.xml'));
  // Content type declared.
  const ct = await outZip.file('[Content_Types].xml')!.async('string');
  assert.match(ct, /externalLink\+xml/);
  // workbook.xml has <externalReferences> and the rId resolves in rels.
  const wbXml = await outZip.file('xl/workbook.xml')!.async('string');
  const refRid = wbXml.match(/<externalReference r:id="([^"]+)"\/>/)?.[1];
  assert.ok(refRid, 'expected an externalReference rId');
  const relsXml = await outZip.file('xl/_rels/workbook.xml.rels')!.async('string');
  assert.match(
    relsXml,
    new RegExp(`Id="${refRid}"[^>]*Target="externalLinks/externalLink1\\.xml"`),
  );
  // <externalReferences> sits before <definedNames>/<calcPr> if present.
  if (wbXml.includes('<calcPr')) {
    assert.ok(
      wbXml.indexOf('<externalReferences>') < wbXml.indexOf('<calcPr'),
      'externalReferences must precede calcPr',
    );
  }

  // ExcelJS can still load the result (it ignores external links, must not throw).
  const reload = new ExcelJS.Workbook();
  await reload.xlsx.load(out);
  assert.ok(reload.getWorksheet('S'));

  // Re-capture round-trips.
  const again = await captureExternalLinksFromBuffer(out);
  assert.deepEqual(again?.order, ['xl/externalLinks/externalLink1.xml']);
});

test('reference order is preserved for multiple links (keeps [N] indices)', async () => {
  // Two links, declared in workbook order 2 then 1, to prove we follow the
  // <externalReferences> order rather than filename order.
  const wb = new ExcelJS.Workbook();
  wb.addWorksheet('S').getCell('A1').value = 1;
  const base = (await wb.xlsx.writeBuffer()) as ArrayBuffer;
  const zip = await JSZip.loadAsync(base);
  for (const n of [1, 2]) {
    zip.file(`xl/externalLinks/externalLink${n}.xml`, LINK_XML);
    zip.file(`xl/externalLinks/_rels/externalLink${n}.xml.rels`, LINK_RELS);
  }
  let ct = await zip.file('[Content_Types].xml')!.async('string');
  for (const n of [1, 2]) {
    ct = ct.replace(
      '</Types>',
      `<Override PartName="/xl/externalLinks/externalLink${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/></Types>`,
    );
  }
  zip.file('[Content_Types].xml', ct);
  let rels = await zip.file('xl/_rels/workbook.xml.rels')!.async('string');
  rels = rels.replace(
    '</Relationships>',
    '<Relationship Id="rId901" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink2.xml"/>' +
      '<Relationship Id="rId902" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/></Relationships>',
  );
  zip.file('xl/_rels/workbook.xml.rels', rels);
  let wbXml = await zip.file('xl/workbook.xml')!.async('string');
  const block =
    '<externalReferences><externalReference r:id="rId901"/><externalReference r:id="rId902"/></externalReferences>';
  wbXml = wbXml.includes('<definedNames')
    ? wbXml.replace('<definedNames', `${block}<definedNames`)
    : wbXml.replace('</workbook>', `${block}</workbook>`);
  zip.file('xl/workbook.xml', wbXml);
  const fixture = await zip.generateAsync({ type: 'arraybuffer' });

  const payload = await captureExternalLinksFromBuffer(fixture);
  // Order follows <externalReferences> (link2 first, then link1) — not filename.
  assert.deepEqual(payload!.order, [
    'xl/externalLinks/externalLink2.xml',
    'xl/externalLinks/externalLink1.xml',
  ]);
});
