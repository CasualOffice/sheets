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
import { waitForUniver } from './_helpers';

/**
 * Smoke test for the perf harness — verifies that the wrapped hot paths
 * actually populate `globalThis.__perf__`. Real load-time benchmarking
 * runs against a large fixture (see docs/LARGE_FILE_PIPELINE.md Stage 7);
 * this file just locks in that the marks fire and are queryable, so the
 * later benchmark tests can rely on them.
 */
test('perf harness records snapshot-save + export-xlsx', async ({ page, baseURL }) => {
  await page.goto(baseURL!);
  await waitForUniver(page);

  // Trigger a save through the export path. workbookDataToXlsx +
  // wb.save() both have timing wrappers; one save should produce both
  // records.
  page.on('download', () => {
    /* let the download finalize so the export-xlsx mark records */
  });
  const downloadPromise = page.waitForEvent('download');
  await page.getByTestId('menubar-file').click();
  // Polish #5 renamed Export → Save as (submenu) with id 'save-as'.
  await page.getByTestId('menu-item-save-as').hover();
  await page.getByTestId('menu-item-save-as-xlsx').click();
  await downloadPromise;

  // Read out the perf buffer.
  const records = await page.evaluate(() => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return ((globalThis as any).__perf__ ?? []) as Array<{ label: string; ms: number }>;
  });

  const labels = records.map((r) => r.label);
  expect(labels, 'expected snapshot-save record').toContain('snapshot-save');
  expect(labels, 'expected export-xlsx record').toContain('export-xlsx');

  const exportMs = records.find((r) => r.label === 'export-xlsx')?.ms ?? 0;
  expect(exportMs).toBeGreaterThan(0);
});
