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
 * Conditional-formatting fidelity. Highlight-cell rules (cellIs numeric
 * comparisons + expression/formula rules, with a fill/font style) were dropped
 * on import. They now bridge to Univer's `SHEET_CONDITIONAL_FORMATTING_PLUGIN`
 * resource and survive a full xlsx round-trip.
 */
declare global {
  interface Window {
    __xlsx?: {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      xlsxToWorkbookData: (buf: ArrayBuffer) => Promise<any>;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      workbookDataToXlsx: (data: any) => Promise<Blob>;
    };
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    __composeCfStyle__?: (row: number, col: number) => any;
  }
}

const CF_RESOURCE = 'SHEET_CONDITIONAL_FORMATTING_PLUGIN';

test('conditional formatting highlight rules survive import + round-trip', async ({ page }) => {
  test.setTimeout(60_000);

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('S');
  for (let r = 1; r <= 5; r++) ws.getCell(`A${r}`).value = r * 50;
  // cellIs: A1:A5 > 100 → red fill.
  ws.addConditionalFormatting({
    ref: 'A1:A5',
    rules: [
      {
        type: 'cellIs',
        operator: 'greaterThan',
        formulae: ['100'],
        priority: 1,
        style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFF0000' } } },
      },
    ],
  });
  // expression: B1:B5 formula rule → bold.
  ws.addConditionalFormatting({
    ref: 'B1:B5',
    rules: [
      {
        type: 'expression',
        formulae: ['MOD(ROW(),2)=0'],
        priority: 1,
        style: { font: { bold: true } },
      },
    ],
  });
  const bytes = Array.from(new Uint8Array((await wb.xlsx.writeBuffer()) as ArrayBuffer));

  await page.goto('/');
  await waitForUniver(page);

  const out = await page.evaluate(
    async ({ buf, resKey }: { buf: number[]; resKey: string }) => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const xlsx = await import(/* @vite-ignore */ '/src/xlsx/index.ts' as any);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const rulesOf = (data: any) => {
        const entry = (data.resources ?? []).find((r: { name: string }) => r.name === resKey);
        if (!entry?.data) return [];
        const parsed = JSON.parse(entry.data);
        const sid = Object.keys(parsed)[0];
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        return (parsed[sid] ?? []).map((x: any) => ({
          subType: x.rule?.subType,
          operator: x.rule?.operator,
          value: x.rule?.value,
          bg: x.rule?.style?.bg?.rgb ?? null,
          bold: x.rule?.style?.bl ?? null,
        }));
      };
      const imported = await xlsx.xlsxToWorkbookData(new Uint8Array(buf).buffer);
      const blob = await xlsx.workbookDataToXlsx(imported);
      const round = await xlsx.xlsxToWorkbookData(await blob.arrayBuffer());
      return { imported: rulesOf(imported), round: rulesOf(round) };
    },
    { buf: bytes, resKey: CF_RESOURCE },
  );

  const expected = [
    { subType: 'number', operator: 'greaterThan', value: 100, bg: '#ff0000', bold: null },
    { subType: 'formula', operator: undefined, value: 'MOD(ROW(),2)=0', bg: null, bold: 1 },
  ];
  expect(out.imported).toEqual(expected);
  expect(out.round).toEqual(expected);
});

test('rank / average / time-period / text CF rules survive import + round-trip', async ({
  page,
}) => {
  test.setTimeout(60_000);

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('S');
  for (let r = 1; r <= 8; r++) {
    ws.getCell(`A${r}`).value = r * 10;
    ws.getCell(`B${r}`).value = `item${r}`;
    ws.getCell(`C${r}`).value = new Date(2026, 0, r);
  }
  const fill = {
    fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFFF00' } },
  } as const;
  ws.addConditionalFormatting({
    ref: 'A1:A8',
    rules: [
      { type: 'top10', rank: 3, percent: false, bottom: false, priority: 1, style: fill },
      { type: 'aboveAverage', aboveAverage: true, priority: 2, style: fill },
      { type: 'aboveAverage', aboveAverage: false, priority: 3, style: fill },
    ],
  });
  ws.addConditionalFormatting({
    ref: 'C1:C8',
    rules: [{ type: 'timePeriod', timePeriod: 'lastWeek', priority: 4, style: fill }],
  });
  ws.addConditionalFormatting({
    ref: 'B1:B8',
    rules: [
      { type: 'containsText', operator: 'containsText', text: 'item', priority: 5, style: fill },
      { type: 'containsText', operator: 'containsBlanks', priority: 6, style: fill },
    ],
  });
  const bytes = Array.from(new Uint8Array((await wb.xlsx.writeBuffer()) as ArrayBuffer));

  await page.goto('/');
  await waitForUniver(page);

  const out = await page.evaluate(
    async ({ buf, resKey }: { buf: number[]; resKey: string }) => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const xlsx = await import(/* @vite-ignore */ '/src/xlsx/index.ts' as any);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const rulesOf = (data: any) => {
        const entry = (data.resources ?? []).find((r: { name: string }) => r.name === resKey);
        if (!entry?.data) return [];
        const parsed = JSON.parse(entry.data);
        // Flatten across all sheets, in rule order.
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        return Object.values(parsed).flatMap((arr: any) =>
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          (arr ?? []).map((x: any) => ({
            subType: x.rule?.subType,
            operator: x.rule?.operator ?? null,
            value: x.rule?.value ?? null,
            isBottom: x.rule?.isBottom ?? null,
            isPercent: x.rule?.isPercent ?? null,
          })),
        );
      };
      const imported = await xlsx.xlsxToWorkbookData(new Uint8Array(buf).buffer);
      const blob = await xlsx.workbookDataToXlsx(imported);
      const round = await xlsx.xlsxToWorkbookData(await blob.arrayBuffer());
      return { imported: rulesOf(imported), round: rulesOf(round) };
    },
    { buf: bytes, resKey: CF_RESOURCE },
  );

  const expected = [
    { subType: 'rank', operator: null, value: 3, isBottom: false, isPercent: false },
    { subType: 'average', operator: 'greaterThan', value: null, isBottom: null, isPercent: null },
    { subType: 'average', operator: 'lessThan', value: null, isBottom: null, isPercent: null },
    { subType: 'timePeriod', operator: 'lastWeek', value: null, isBottom: null, isPercent: null },
    { subType: 'text', operator: 'containsText', value: 'item', isBottom: null, isPercent: null },
    {
      subType: 'text',
      operator: 'containsBlanks',
      value: null,
      isBottom: null,
      isPercent: null,
    },
  ];
  // Order is grouped by sheet-ref insertion; compare as sets to stay robust.
  expect([...out.imported].sort(byKey)).toEqual([...expected].sort(byKey));
  expect([...out.round].sort(byKey)).toEqual([...expected].sort(byKey));
});

// Stable sort key for order-independent rule comparison.
function byKey(a: { subType: string; operator: string | null; value: unknown }, b: typeof a) {
  return `${a.subType}:${a.operator}:${JSON.stringify(a.value)}`.localeCompare(
    `${b.subType}:${b.operator}:${JSON.stringify(b.value)}`,
  );
}

/**
 * Render-on-import: a rule that round-trips in the resource is not enough — it
 * must actually MATCH + PAINT when the file opens. The CF number-rule evaluator
 * keys off the cell's `t` (CellValueType) directly, so an imported numeric cell
 * that lacked `t` never matched a `cellIs` rule — the highlight stayed blank no
 * matter how long you waited or how much you scrolled. parse-impl now tags
 * imported cells with their value type. We read the result through the CF
 * service's `composeStyle` (the highlight is canvas-drawn, not in cell data),
 * asserting a matching cell composes the red fill while a non-matching cell
 * composes nothing — immediately on open, without any interaction.
 */
test('imported conditional formatting paints on open without interaction', async ({ page }) => {
  test.setTimeout(60_000);

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('S');
  // A1..A5 = 50,100,150,200,250 → only A3:A5 (>100) should highlight.
  for (let r = 1; r <= 5; r++) ws.getCell(`A${r}`).value = r * 50;
  ws.addConditionalFormatting({
    ref: 'A1:A5',
    rules: [
      {
        type: 'cellIs',
        operator: 'greaterThan',
        formulae: ['100'],
        priority: 1,
        style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFF0000' } } },
      },
    ],
  });
  const bytes = Array.from(new Uint8Array((await wb.xlsx.writeBuffer()) as ArrayBuffer));

  await page.goto('/');
  await waitForUniver(page);

  // Write the fixture and open it through the real File → Open picker so the
  // workbook-swap path (UniverSheet revision effect) runs, exactly as a user's
  // open would.
  const fixture = '/tmp/casual-sheets-cf-render.xlsx';
  const fs = await import('node:fs');
  fs.writeFileSync(fixture, Buffer.from(bytes));

  await page.getByTestId('menubar-file').click();
  const [chooser] = await Promise.all([
    page.waitForEvent('filechooser'),
    page.getByTestId('menu-item-open').click(),
  ]);
  await chooser.setFiles(fixture);

  // The matching cell A5 (row 4, col 0, value 250 > 100) must compose the red
  // fill — poll because the CF recompute the pump triggers is async, but assert
  // WITHOUT any scroll/click that would itself force a recompute.
  await expect
    .poll(
      async () =>
        page.evaluate(() => {
          const s = window.__composeCfStyle__?.(4, 0);
          return (s?.style?.bg?.rgb ?? null) as string | null;
        }),
      { timeout: 10_000, message: 'A5 should compose the imported red CF fill on open' },
    )
    .toBe('#ff0000');

  // The non-matching cell A1 (row 0, col 0, value 50) composes no CF style.
  const a1 = await page.evaluate(() => {
    const s = window.__composeCfStyle__?.(0, 0);
    return s?.style?.bg?.rgb ?? null;
  });
  expect(a1).toBeNull();
});

/**
 * Render-on-import for the precomputing rule types (rank / average / text):
 * these scan the whole range before deciding a cell, a different evaluator path
 * than the per-cell `number` rule above. A top-3 rank over A1:A8 (10..80) must
 * paint A8 (=80, in the top 3) and leave A1 (=10) blank — immediately on open.
 */
test('imported rank conditional formatting paints on open', async ({ page }) => {
  test.setTimeout(60_000);

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('S');
  for (let r = 1; r <= 8; r++) ws.getCell(`A${r}`).value = r * 10;
  ws.addConditionalFormatting({
    ref: 'A1:A8',
    rules: [
      {
        type: 'top10',
        rank: 3,
        percent: false,
        bottom: false,
        priority: 1,
        style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FF00FF00' } } },
      },
    ],
  });
  const bytes = Array.from(new Uint8Array((await wb.xlsx.writeBuffer()) as ArrayBuffer));

  await page.goto('/');
  await waitForUniver(page);

  const fixture = '/tmp/casual-sheets-cf-rank.xlsx';
  const fs = await import('node:fs');
  fs.writeFileSync(fixture, Buffer.from(bytes));

  await page.getByTestId('menubar-file').click();
  const [chooser] = await Promise.all([
    page.waitForEvent('filechooser'),
    page.getByTestId('menu-item-open').click(),
  ]);
  await chooser.setFiles(fixture);

  // A8 (row 7, col 0) = 80 is in the top 3 → green fill, no interaction.
  await expect
    .poll(
      async () =>
        page.evaluate(() => {
          const s = window.__composeCfStyle__?.(7, 0);
          return (s?.style?.bg?.rgb ?? null) as string | null;
        }),
      { timeout: 10_000, message: 'A8 should compose the top-3 rank fill on open' },
    )
    .toBe('#00ff00');

  // A1 (=10) is not in the top 3 → no CF style.
  const a1 = await page.evaluate(() => {
    const s = window.__composeCfStyle__?.(0, 0);
    return s?.style?.bg?.rgb ?? null;
  });
  expect(a1).toBeNull();
});

test('color-scale CF survives round-trip and paints a value-mapped gradient on open', async ({
  page,
}) => {
  test.setTimeout(60_000);

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('S');
  for (let r = 1; r <= 8; r++) ws.getCell(`A${r}`).value = r * 10;
  // 3-colour scale: min = red, midpoint = yellow, max = green.
  ws.addConditionalFormatting({
    ref: 'A1:A8',
    rules: [
      {
        type: 'colorScale',
        priority: 1,
        cfvo: [{ type: 'min' }, { type: 'percentile', value: 50 }, { type: 'max' }],
        color: [{ argb: 'FFF8696B' }, { argb: 'FFFFEB84' }, { argb: 'FF63BE7B' }],
      },
    ],
  });
  const bytes = Array.from(new Uint8Array((await wb.xlsx.writeBuffer()) as ArrayBuffer));

  await page.goto('/');
  await waitForUniver(page);

  // First assert the resource round-trips (import → export → re-import).
  const round = await page.evaluate(
    async ({ buf, resKey }: { buf: number[]; resKey: string }) => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const xlsx = await import(/* @vite-ignore */ '/src/xlsx/index.ts' as any);
      const imported = await xlsx.xlsxToWorkbookData(new Uint8Array(buf).buffer);
      const blob = await xlsx.workbookDataToXlsx(imported);
      const re = await xlsx.xlsxToWorkbookData(await blob.arrayBuffer());
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const cfg = (data: any) => {
        const entry = (data.resources ?? []).find((r: { name: string }) => r.name === resKey);
        const parsed = JSON.parse(entry.data);
        return parsed[Object.keys(parsed)[0]][0].rule;
      };
      return { type: cfg(re).type, stops: cfg(re).config.length };
    },
    { buf: bytes, resKey: CF_RESOURCE },
  );
  expect(round).toEqual({ type: 'colorScale', stops: 3 });

  // Then open it and assert the gradient paints on the canvas without interaction:
  // the min cell composes the red endpoint, the max cell the green endpoint.
  const fixture = '/tmp/casual-sheets-cf-colorscale.xlsx';
  const fs = await import('node:fs');
  fs.writeFileSync(fixture, Buffer.from(bytes));

  await page.getByTestId('menubar-file').click();
  const [chooser] = await Promise.all([
    page.waitForEvent('filechooser'),
    page.getByTestId('menu-item-open').click(),
  ]);
  await chooser.setFiles(fixture);

  // The gradient endpoints render as the exact stop colours. ColorKit emits
  // interpolated fills as `rgb(r,g,b)` strings, so compare in that form:
  // #f8696b = rgb(248,105,107), #63be7b = rgb(99,190,123).
  // A1 (=10, the range min) → red endpoint.
  await expect
    .poll(
      async () =>
        page.evaluate(() => {
          const s = window.__composeCfStyle__?.(0, 0);
          return (s?.style?.bg?.rgb ?? null) as string | null;
        }),
      { timeout: 10_000, message: 'A1 should compose the color-scale min (red) on open' },
    )
    .toBe('rgb(248,105,107)');

  // A8 (=80, the range max) → green endpoint.
  const a8 = await page.evaluate(() => {
    const s = window.__composeCfStyle__?.(7, 0);
    return s?.style?.bg?.rgb ?? null;
  });
  expect(a8).toBe('rgb(99,190,123)');
});

test('icon-set CF survives round-trip and paints the right icon per band on open', async ({
  page,
}) => {
  test.setTimeout(60_000);

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('S');
  for (let r = 1; r <= 8; r++) ws.getCell(`A${r}`).value = r * 10;
  // 3 traffic lights with absolute thresholds: >=67 top, >=33 mid, else bottom.
  ws.addConditionalFormatting({
    ref: 'A1:A8',
    rules: [
      {
        type: 'iconSet',
        priority: 1,
        iconSet: '3TrafficLights1',
        cfvo: [
          { type: 'num', value: 0 },
          { type: 'num', value: 33 },
          { type: 'num', value: 67 },
        ],
      },
    ],
  });
  const bytes = Array.from(new Uint8Array((await wb.xlsx.writeBuffer()) as ArrayBuffer));

  await page.goto('/');
  await waitForUniver(page);

  // Resource round-trips.
  const round = await page.evaluate(
    async ({ buf, resKey }: { buf: number[]; resKey: string }) => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const xlsx = await import(/* @vite-ignore */ '/src/xlsx/index.ts' as any);
      const imported = await xlsx.xlsxToWorkbookData(new Uint8Array(buf).buffer);
      const blob = await xlsx.workbookDataToXlsx(imported);
      const re = await xlsx.xlsxToWorkbookData(await blob.arrayBuffer());
      const entry = (re.resources ?? []).find((r: { name: string }) => r.name === resKey);
      const parsed = JSON.parse(entry.data);
      const rule = parsed[Object.keys(parsed)[0]][0].rule;
      return { type: rule.type, bands: rule.config.length, topIcon: rule.config[0].iconId };
    },
    { buf: bytes, resKey: CF_RESOURCE },
  );
  expect(round).toEqual({ type: 'iconSet', bands: 3, topIcon: '0' });

  // Open and assert the painted icon per band, no interaction.
  const fixture = '/tmp/casual-sheets-cf-iconset.xlsx';
  const fs = await import('node:fs');
  fs.writeFileSync(fixture, Buffer.from(bytes));

  await page.getByTestId('menubar-file').click();
  const [chooser] = await Promise.all([
    page.waitForEvent('filechooser'),
    page.getByTestId('menu-item-open').click(),
  ]);
  await chooser.setFiles(fixture);

  // A8 (=80, >=67) → top icon (iconId '0').
  await expect
    .poll(
      async () =>
        page.evaluate(() => {
          const s = window.__composeCfStyle__?.(7, 0);
          return (s?.iconSet?.iconId ?? null) as string | null;
        }),
      { timeout: 10_000, message: 'A8 should compose the top icon on open' },
    )
    .toBe('0');

  // A1 (=10, < 33) → bottom icon (iconId '2'); A5 (=50, [33,67)) → mid (iconId '1').
  const [a1, a5] = await page.evaluate(() => [
    window.__composeCfStyle__?.(0, 0)?.iconSet?.iconId ?? null,
    window.__composeCfStyle__?.(4, 0)?.iconSet?.iconId ?? null,
  ]);
  expect(a1).toBe('2');
  expect(a5).toBe('1');
});

test('data-bar CF (positive colour recovered from raw XML) paints on open', async ({ page }) => {
  test.setTimeout(60_000);

  // ExcelJS writes a broken <color auto="1"/> for data bars, so author the
  // fixture and patch in a real <color rgb> — exactly what a real Excel file
  // carries and what the importer recovers from raw XML.
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('S');
  for (let r = 1; r <= 5; r++) ws.getCell(`A${r}`).value = r * 10;
  ws.addConditionalFormatting({
    ref: 'A1:A5',
    rules: [{ type: 'dataBar', cfvo: [{ type: 'min' }, { type: 'max' }], priority: 1 }],
  });
  const raw = (await wb.xlsx.writeBuffer()) as ArrayBuffer;
  const zip = await JSZip.loadAsync(raw);
  let sheetXml = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
  sheetXml = sheetXml.replace('<color auto="1"/>', '<color rgb="FF638EC6"/>');
  zip.file('xl/worksheets/sheet1.xml', sheetXml);
  const bytes = Array.from(
    new Uint8Array((await zip.generateAsync({ type: 'arraybuffer' })) as ArrayBuffer),
  );

  await page.goto('/');
  await waitForUniver(page);

  const fixture = '/tmp/casual-sheets-cf-databar.xlsx';
  const fs = await import('node:fs');
  fs.writeFileSync(fixture, Buffer.from(bytes));

  await page.getByTestId('menubar-file').click();
  const [chooser] = await Promise.all([
    page.waitForEvent('filechooser'),
    page.getByTestId('menu-item-open').click(),
  ]);
  await chooser.setFiles(fixture);

  // A3 (row 2) is inside the data-bar range → composes a dataBar render entry.
  await expect
    .poll(
      async () =>
        page.evaluate(() => {
          const s = window.__composeCfStyle__?.(2, 0);
          return s?.dataBar ? true : false;
        }),
      { timeout: 10_000, message: 'A3 should compose a data bar on open' },
    )
    .toBe(true);

  // A cell outside the range composes nothing.
  const outside = await page.evaluate(() => {
    const s = window.__composeCfStyle__?.(2, 5); // F3
    return s?.dataBar ? true : false;
  });
  expect(outside).toBe(false);
});

test('duplicate-values CF (raw-XML rule + dxf) paints duplicate cells on open', async ({
  page,
}) => {
  test.setTimeout(60_000);

  // A1=10, A2=20, A3=10 (dup), A4=30, A5=20 (dup) → 10 and 20 are duplicates.
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('S');
  [10, 20, 10, 30, 20].forEach((v, i) => (ws.getCell(`A${i + 1}`).value = v));
  const raw = (await wb.xlsx.writeBuffer()) as ArrayBuffer;

  // Inject a duplicateValues cfRule + its green dxf (ExcelJS writes neither).
  // ExcelJS emits an empty <dxfs count="0"/>; populate it in place — adding a
  // second <dxfs> element would crash ExcelJS's loader on reconcile.
  const zip = await JSZip.loadAsync(raw);
  let styles = await zip.file('xl/styles.xml')!.async('string');
  styles = styles.replace(
    /<dxfs count="0"\/>/,
    '<dxfs count="1"><dxf><fill><patternFill patternType="solid"><bgColor rgb="FF00FF00"/></patternFill></fill></dxf></dxfs>',
  );
  zip.file('xl/styles.xml', styles);
  let sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
  sheet = sheet.replace(
    '<pageMargins',
    '<conditionalFormatting sqref="A1:A5"><cfRule type="duplicateValues" dxfId="0" priority="1"/></conditionalFormatting><pageMargins',
  );
  zip.file('xl/worksheets/sheet1.xml', sheet);
  const bytes = Array.from(
    new Uint8Array((await zip.generateAsync({ type: 'arraybuffer' })) as ArrayBuffer),
  );

  await page.goto('/');
  await waitForUniver(page);

  const fixture = '/tmp/casual-sheets-cf-dup.xlsx';
  const fs = await import('node:fs');
  fs.writeFileSync(fixture, Buffer.from(bytes));

  await page.getByTestId('menubar-file').click();
  const [chooser] = await Promise.all([
    page.waitForEvent('filechooser'),
    page.getByTestId('menu-item-open').click(),
  ]);
  await chooser.setFiles(fixture);

  // A1 (=10, duplicated) → green fill on open, no interaction.
  await expect
    .poll(
      async () =>
        page.evaluate(() => {
          const s = window.__composeCfStyle__?.(0, 0);
          return (s?.style?.bg?.rgb ?? null) as string | null;
        }),
      { timeout: 10_000, message: 'A1 (duplicate) should compose the green fill on open' },
    )
    .toBe('#00ff00');

  // A4 (=30, unique) → no CF style.
  const a4 = await page.evaluate(() => {
    const s = window.__composeCfStyle__?.(3, 0);
    return s?.style?.bg?.rgb ?? null;
  });
  expect(a4).toBeNull();
});

test('beginsWith text CF paints matching cells on open', async ({ page }) => {
  test.setTimeout(60_000);

  // A1=PROJ, A2=other, A3=PRINT → "PR" prefix matches A1 and A3.
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('S');
  ['PROJ', 'other', 'PRINT'].forEach((v, i) => (ws.getCell(`A${i + 1}`).value = v));
  ws.addConditionalFormatting({
    ref: 'A1:A3',
    rules: [
      {
        type: 'containsText',
        operator: 'beginsWith',
        formulae: ['LEFT(A1,LEN("PR"))="PR"'],
        priority: 1,
        style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FF00FF00' } } },
      },
    ],
  });
  const bytes = Array.from(new Uint8Array((await wb.xlsx.writeBuffer()) as ArrayBuffer));

  await page.goto('/');
  await waitForUniver(page);

  const fixture = '/tmp/casual-sheets-cf-beginswith.xlsx';
  const fs = await import('node:fs');
  fs.writeFileSync(fixture, Buffer.from(bytes));

  await page.getByTestId('menubar-file').click();
  const [chooser] = await Promise.all([
    page.waitForEvent('filechooser'),
    page.getByTestId('menu-item-open').click(),
  ]);
  await chooser.setFiles(fixture);

  // A1 (PROJ, begins "PR") → green on open.
  await expect
    .poll(
      async () =>
        page.evaluate(() => {
          const s = window.__composeCfStyle__?.(0, 0);
          return (s?.style?.bg?.rgb ?? null) as string | null;
        }),
      { timeout: 10_000, message: 'A1 (begins with PR) should compose green on open' },
    )
    .toBe('#00ff00');

  // A2 (other) → no CF style.
  const a2 = await page.evaluate(() => window.__composeCfStyle__?.(1, 0)?.style?.bg?.rgb ?? null);
  expect(a2).toBeNull();
});
