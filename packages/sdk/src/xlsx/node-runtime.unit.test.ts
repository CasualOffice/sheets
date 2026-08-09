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
import { xlsxToWorkbookData } from './import';
import { workbookDataToXlsx } from './export';

test('imports and exports xlsx in Node without browser Worker globals', async () => {
  assert.equal(typeof window, 'undefined');

  const source = new ExcelJS.Workbook();
  source.title = 'Node round-trip';
  const sheet = source.addWorksheet('Data');
  sheet.getCell('A1').value = 'preserved';
  const input = await source.xlsx.writeBuffer();
  const inputBuffer = input.buffer.slice(
    input.byteOffset,
    input.byteOffset + input.byteLength,
  ) as ArrayBuffer;

  const snapshot = await xlsxToWorkbookData(inputBuffer);
  assert.equal(snapshot.sheets[snapshot.sheetOrder[0]]?.cellData?.[0]?.[0]?.v, 'preserved');
  assert.equal(snapshot.appVersion, '0.25.0');
  snapshot.resources = [{ name: 'HOST_OPAQUE_RESOURCE', data: 'opaque-payload' }];

  const output = await workbookDataToXlsx(snapshot);
  assert.ok(output instanceof Blob);
  assert.equal(output.type, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  const outputBuffer = await output.arrayBuffer();

  const reopened = new ExcelJS.Workbook();
  await reopened.xlsx.load(outputBuffer);
  assert.equal(reopened.getWorksheet('Data')?.getCell('A1').value, 'preserved');

  const reparsed = await xlsxToWorkbookData(outputBuffer);
  assert.deepEqual(
    reparsed.resources?.find((resource) => resource.name === 'HOST_OPAQUE_RESOURCE'),
    {
      name: 'HOST_OPAQUE_RESOURCE',
      data: 'opaque-payload',
    },
  );
});
