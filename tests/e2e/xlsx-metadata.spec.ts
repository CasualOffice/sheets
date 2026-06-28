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

import { expect, test, type Page } from '@playwright/test';
import { waitForUniver } from './_helpers';

/**
 * File metadata (title / subject / author / tags / category / description /
 * company / manager / created / modified) lives in `custom.properties`
 * on our snapshot. Without explicit mapping it never reaches the xlsx
 * Core/App properties (docProps/core.xml + docProps/app.xml) — and
 * never comes back on re-open. This spec locks the round-trip.
 */

async function exposeConverters(page: Page) {
  await page.evaluate(async () => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const mod = await import(/* @vite-ignore */ '/src/xlsx/index.ts' as any);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (window as any).__xlsx = mod;
  });
}

test('file metadata round-trips through xlsx save/load', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);
  await exposeConverters(page);

  const reloadedProps = await page.evaluate(async () => {
    const original = {
      id: 'wb-meta',
      rev: 1,
      name: 'My Workbook',
      appVersion: '0.22.1',
      locale: 1,
      styles: {},
      sheetOrder: ['s1'],
      sheets: {
        s1: { id: 's1', name: 'Sheet1', cellData: {}, rowCount: 100, columnCount: 26 },
      },
      // The shape the Properties dialog writes via custom.properties.
      custom: {
        properties: {
          title: 'Q3 Revenue',
          subject: 'Finance',
          author: 'Sachin S.',
          tags: 'Q3,2026,Revenue',
          category: 'Quarterly',
          description: 'Q3 2026 revenue rollup with regional breakdown.',
          company: 'Acme Corp',
          manager: 'A. Manager',
          createdAt: '2026-05-01T09:00:00.000Z',
          // modifiedAt is bumped to "now" by the exporter, not preserved.
        },
      },
    };
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const xlsx = (window as any).__xlsx!;
    const blob = await xlsx.workbookDataToXlsx(original);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const reloaded: any = await xlsx.xlsxToWorkbookData(await blob.arrayBuffer());
    return reloaded.custom?.properties ?? null;
  });

  expect(reloadedProps).not.toBeNull();
  expect(reloadedProps.title).toBe('Q3 Revenue');
  expect(reloadedProps.subject).toBe('Finance');
  expect(reloadedProps.author).toBe('Sachin S.');
  expect(reloadedProps.tags).toBe('Q3,2026,Revenue');
  expect(reloadedProps.category).toBe('Quarterly');
  expect(reloadedProps.description).toBe('Q3 2026 revenue rollup with regional breakdown.');
  expect(reloadedProps.company).toBe('Acme Corp');
  expect(reloadedProps.manager).toBe('A. Manager');
  expect(reloadedProps.createdAt).toBe('2026-05-01T09:00:00.000Z');
  // modifiedAt is freshly stamped by the exporter; just verify it's
  // present and parseable as a date.
  expect(typeof reloadedProps.modifiedAt).toBe('string');
  expect(Number.isFinite(new Date(reloadedProps.modifiedAt).getTime())).toBe(true);
});
