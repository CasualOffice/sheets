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
import { waitForUniver } from './_helpers';

/**
 * Embedded-image round-trip. Univer has no drawing model, so ExcelJS rebuilds
 * the exported workbook without any image — opening an xlsx and saving it used
 * to silently drop every embedded picture (tracker #192, the power-user audit's
 * #1 data-loss blocker). The drawing-passthrough layer captures xl/media +
 * xl/drawings at parse time and re-injects them at export. This drives the real
 * in-app pipeline (xlsxToWorkbookData → workbookDataToXlsx) and asserts the
 * image still loads after the round-trip.
 */

// 1×1 transparent PNG.
const PNG_1X1_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

test('an embedded image survives an open → save round-trip', async ({ page }) => {
  test.setTimeout(60_000);

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Pics');
  ws.getCell('A1').value = 'logo below';
  const imageId = wb.addImage({ base64: PNG_1X1_BASE64, extension: 'png' });
  ws.addImage(imageId, 'B2:D6');
  const bytes = Array.from(new Uint8Array((await wb.xlsx.writeBuffer()) as ArrayBuffer));

  await page.goto('/');
  await waitForUniver(page);

  // Run the actual SDK import → export pipeline in the page, return the bytes.
  const outBytes = await page.evaluate(async (buf: number[]) => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const xlsx = await import(/* @vite-ignore */ '/src/xlsx/index.ts' as any);
    const imported = await xlsx.xlsxToWorkbookData(new Uint8Array(buf).buffer);
    const blob = await xlsx.workbookDataToXlsx(imported);
    return Array.from(new Uint8Array(await blob.arrayBuffer()));
  }, bytes);

  // The exported file must still contain the image.
  const reload = new ExcelJS.Workbook();
  await reload.xlsx.load(new Uint8Array(outBytes).buffer as ArrayBuffer);
  const images = reload.getWorksheet('Pics')!.getImages();
  expect(images.length, 'image should survive the round-trip').toBe(1);

  // The media bytes are intact (a real PNG, non-empty).
  const media = reload.model.media?.find((m) => m.type === 'image');
  expect(media, 'media payload should be present').toBeTruthy();
  expect((media!.buffer as Buffer).byteLength).toBeGreaterThan(0);
});
