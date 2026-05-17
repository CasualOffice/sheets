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
  await page.getByTestId('menu-item-export').hover();
  await page.getByTestId('menu-item-export-xlsx').click();
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
