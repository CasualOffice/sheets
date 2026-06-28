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

import { expect, test } from '@playwright/test';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { waitForUniver } from './_helpers';

/**
 * External-workbook-link round-trip. A formula like `=[1]Sheet1!A1` references
 * another workbook via `<externalReferences>` → `xl/externalLinks/**`. ExcelJS
 * has no external-link model, so it rebuilt the export without those parts —
 * the `[1]` index dangled and the formula resolved to `#REF!` (silent
 * corruption, tracker #192). The external-link passthrough captures and
 * re-injects them. This drives the real in-app pipeline and asserts the parts +
 * workbook references survive.
 */

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

async function bufferWithExternalLink(): Promise<ArrayBuffer> {
  const wb = new ExcelJS.Workbook();
  wb.addWorksheet('S').getCell('A1').value = { formula: '[1]Sheet1!A1', result: 42 };
  const base = (await wb.xlsx.writeBuffer()) as ArrayBuffer;

  const zip = await JSZip.loadAsync(base);
  zip.file('xl/externalLinks/externalLink1.xml', LINK_XML);
  zip.file('xl/externalLinks/_rels/externalLink1.xml.rels', LINK_RELS);
  let ct = await zip.file('[Content_Types].xml')!.async('string');
  ct = ct.replace(
    '</Types>',
    '<Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/></Types>',
  );
  zip.file('[Content_Types].xml', ct);
  let rels = await zip.file('xl/_rels/workbook.xml.rels')!.async('string');
  rels = rels.replace(
    '</Relationships>',
    '<Relationship Id="rId900" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/></Relationships>',
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

test('external-workbook links survive an open → save round-trip', async ({ page }) => {
  test.setTimeout(60_000);

  const bytes = Array.from(new Uint8Array(await bufferWithExternalLink()));

  await page.goto('/');
  await waitForUniver(page);

  const outBytes = await page.evaluate(async (buf: number[]) => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const xlsx = await import(/* @vite-ignore */ '/src/xlsx/index.ts' as any);
    const imported = await xlsx.xlsxToWorkbookData(new Uint8Array(buf).buffer);
    const blob = await xlsx.workbookDataToXlsx(imported);
    return Array.from(new Uint8Array(await blob.arrayBuffer()));
  }, bytes);

  const outZip = await JSZip.loadAsync(new Uint8Array(outBytes).buffer as ArrayBuffer);
  // The external-link part survived.
  expect(
    outZip.file('xl/externalLinks/externalLink1.xml'),
    'externalLink part should survive',
  ).toBeTruthy();
  // workbook.xml references it, and the rId resolves in rels.
  const wbXml = await outZip.file('xl/workbook.xml')!.async('string');
  const refRid = wbXml.match(/<externalReference r:id="([^"]+)"\/>/)?.[1];
  expect(refRid, 'workbook should declare an externalReference').toBeTruthy();
  const relsXml = await outZip.file('xl/_rels/workbook.xml.rels')!.async('string');
  expect(relsXml).toMatch(
    new RegExp(`Id="${refRid}"[^>]*Target="externalLinks/externalLink1\\.xml"`),
  );
});
