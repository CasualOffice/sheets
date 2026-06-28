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
 * Threaded-comment round-trip. Modern Excel comments carry real authors,
 * timestamps, and reply chains in `xl/threadedComments/**` + `xl/persons/**`.
 * ExcelJS only models the legacy note, so our bridge collapsed every thread to
 * one note authored "imported" and dropped the replies/authors on save
 * (tracker #192). The threaded-comment passthrough captures and re-injects the
 * threaded layer. This drives the real in-app pipeline and asserts the authors
 * + reply chain survive.
 */

const PERSON_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments" ' +
  'xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
  '<person displayName="Ada Lovelace" id="{AAAA0000-0000-0000-0000-000000000001}" userId="ada" providerId="None"/>' +
  '<person displayName="Bob Stone" id="{BBBB0000-0000-0000-0000-000000000002}" userId="bob" providerId="None"/>' +
  '</personList>';

const THREAD_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">' +
  '<threadedComment ref="B2" dT="2026-01-01T10:00:00Z" personId="{AAAA0000-0000-0000-0000-000000000001}" id="{C100-0000-0000-0000-000000000001}"><text>Looks off here</text></threadedComment>' +
  '<threadedComment ref="B2" dT="2026-01-01T11:00:00Z" personId="{BBBB0000-0000-0000-0000-000000000002}" id="{C100-0000-0000-0000-000000000002}" parentId="{C100-0000-0000-0000-000000000001}"><text>Agreed, fixing</text></threadedComment>' +
  '</ThreadedComments>';

async function bufferWithThread(): Promise<ArrayBuffer> {
  const wb = new ExcelJS.Workbook();
  wb.addWorksheet('S').getCell('B2').note = 'Looks off here';
  const base = (await wb.xlsx.writeBuffer()) as ArrayBuffer;

  const zip = await JSZip.loadAsync(base);
  zip.file('xl/persons/person.xml', PERSON_XML);
  zip.file('xl/threadedComments/threadedComment1.xml', THREAD_XML);
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

test('threaded comments (authors + replies) survive an open → save round-trip', async ({
  page,
}) => {
  test.setTimeout(60_000);

  const bytes = Array.from(new Uint8Array(await bufferWithThread()));

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
  // The threaded layer survived with authors + reply chain.
  const thread = outZip.file('xl/threadedComments/threadedComment1.xml');
  expect(thread, 'threadedComment part should survive').toBeTruthy();
  const threadXml = await thread!.async('string');
  expect(threadXml).toMatch(/parentId=/); // reply preserved
  const persons = await outZip.file('xl/persons/person.xml')!.async('string');
  expect(persons).toMatch(/Ada Lovelace/);
  expect(persons).toMatch(/Bob Stone/);
  // sheet→threadedComment relationship re-created.
  const sheetRels = await outZip.file('xl/worksheets/_rels/sheet1.xml.rels')!.async('string');
  expect(sheetRels).toMatch(/relationships\/threadedComment"/);
});
