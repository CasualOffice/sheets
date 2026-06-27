import { strict as assert } from 'node:assert';
import { test } from 'node:test';
import ExcelJS from 'exceljs';

import {
  applyConditionalFormattingToXlsxWorksheet,
  CONDITIONAL_FORMATTING_RESOURCE,
  readConditionalFormattingFromSnapshot,
  readConditionalFormattingFromXlsx,
} from './conditional-formatting-resource.js';

const FILL = {
  fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFF0000' } },
} as const;

// Identity sheet-id mapping keyed by the worksheet's ExcelJS id.
const sheetId = (excelId: number) => `s-${excelId}`;

function buildWorkbook(rules: object[]): ExcelJS.Workbook {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('S');
  for (let r = 1; r <= 8; r++) {
    ws.getCell(`A${r}`).value = r * 10;
    ws.getCell(`B${r}`).value = `item${r}`;
  }
  ws.addConditionalFormatting({
    ref: 'A1:A8',
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    rules: rules as any,
  });
  return wb;
}

/** Build → import → export → reload → import again, returning the rule shapes
 *  after the full bridge round-trip (the only path that exercises ExcelJS's
 *  own read/write quirks). */
async function roundTrip(rules: object[]) {
  const wb = buildWorkbook(rules);
  const firstPass = readConditionalFormattingFromXlsx(wb, sheetId);
  const synthRules = Object.values(firstPass)[0] ?? [];

  const outWb = new ExcelJS.Workbook();
  const outWs = outWb.addWorksheet('S');
  applyConditionalFormattingToXlsxWorksheet(outWs, synthRules);

  const buf = await outWb.xlsx.writeBuffer();
  const reloaded = new ExcelJS.Workbook();
  await reloaded.xlsx.load(buf as Buffer);
  const secondPass = readConditionalFormattingFromXlsx(reloaded, sheetId);
  return (Object.values(secondPass)[0] ?? []).map((r) => r.rule);
}

test('cellIs numeric rule round-trips operator + value + style', async () => {
  const out = await roundTrip([
    { type: 'cellIs', operator: 'greaterThan', formulae: ['100'], priority: 1, style: FILL },
    { type: 'cellIs', operator: 'between', formulae: ['10', '50'], priority: 2, style: FILL },
  ]);
  assert.deepEqual(out[0], {
    type: 'highlightCell',
    subType: 'number',
    operator: 'greaterThan',
    value: 100,
    style: { bg: { rgb: '#ff0000' } },
  });
  assert.deepEqual(out[1].subType, 'number');
  assert.deepEqual((out[1] as { value: unknown }).value, [10, 50]);
});

test('top10 maps to rank with isBottom / isPercent / value', async () => {
  const out = await roundTrip([
    { type: 'top10', rank: 3, percent: false, bottom: false, priority: 1, style: FILL },
    { type: 'top10', rank: 25, percent: true, bottom: true, priority: 2, style: FILL },
  ]);
  assert.equal(out[0].subType, 'rank');
  assert.deepEqual(out[0], {
    type: 'highlightCell',
    subType: 'rank',
    isBottom: false,
    isPercent: false,
    value: 3,
    style: { bg: { rgb: '#ff0000' } },
  });
  assert.deepEqual(out[1], {
    type: 'highlightCell',
    subType: 'rank',
    isBottom: true,
    isPercent: true,
    value: 25,
    style: { bg: { rgb: '#ff0000' } },
  });
});

test('aboveAverage maps to average operator (above→greaterThan, below→lessThan)', async () => {
  const out = await roundTrip([
    { type: 'aboveAverage', aboveAverage: true, priority: 1, style: FILL },
    { type: 'aboveAverage', aboveAverage: false, priority: 2, style: FILL },
  ]);
  assert.equal(out[0].subType, 'average');
  assert.equal((out[0] as { operator: string }).operator, 'greaterThan');
  assert.equal((out[1] as { operator: string }).operator, 'lessThan');
});

test('timePeriod maps operator verbatim', async () => {
  const out = await roundTrip([
    { type: 'timePeriod', timePeriod: 'lastWeek', priority: 1, style: FILL },
    { type: 'timePeriod', timePeriod: 'thisMonth', priority: 2, style: FILL },
  ]);
  assert.equal(out[0].subType, 'timePeriod');
  assert.equal((out[0] as { operator: string }).operator, 'lastWeek');
  assert.equal((out[1] as { operator: string }).operator, 'thisMonth');
});

test('containsText recovers search value from formula; blanks/errors are value-less', async () => {
  const out = await roundTrip([
    { type: 'containsText', operator: 'containsText', text: 'item', priority: 1, style: FILL },
    { type: 'containsText', operator: 'containsBlanks', priority: 2, style: FILL },
    { type: 'containsText', operator: 'containsErrors', priority: 3, style: FILL },
  ]);
  assert.deepEqual(out[0], {
    type: 'highlightCell',
    subType: 'text',
    operator: 'containsText',
    value: 'item',
    style: { bg: { rgb: '#ff0000' } },
  });
  assert.equal(out[1].subType, 'text');
  assert.equal((out[1] as { operator: string }).operator, 'containsBlanks');
  assert.equal((out[1] as { value?: string }).value, undefined);
  assert.equal((out[2] as { operator: string }).operator, 'containsErrors');
});

test('unmappable rules (beginsWith, duplicateValues) are skipped, not corrupted', async () => {
  // ExcelJS drops these on round-trip; the bridge must not emit a broken rule.
  const out = await roundTrip([
    { type: 'containsText', operator: 'beginsWith', text: 'it', priority: 1, style: FILL },
    { type: 'duplicateValues', priority: 2, style: FILL },
    // A mappable rule alongside them still survives.
    { type: 'top10', rank: 2, percent: false, bottom: false, priority: 3, style: FILL },
  ]);
  assert.deepEqual(
    out.map((r) => r.subType),
    ['rank'],
  );
});

test('colorScale maps cfvo + color to ordered gradient stops', async () => {
  const out = await roundTrip([
    {
      type: 'colorScale',
      priority: 1,
      cfvo: [{ type: 'min' }, { type: 'percentile', value: 50 }, { type: 'max' }],
      color: [{ argb: 'FFF8696B' }, { argb: 'FFFFEB84' }, { argb: 'FF63BE7B' }],
    },
  ]);
  assert.equal(out.length, 1);
  assert.equal(out[0].type, 'colorScale');
  assert.deepEqual((out[0] as { config: unknown }).config, [
    { index: 0, color: '#f8696b', value: { type: 'min' } },
    { index: 1, color: '#ffeb84', value: { type: 'percentile', value: 50 } },
    { index: 2, color: '#63be7b', value: { type: 'max' } },
  ]);
});

test('colorScale with a formula threshold is dropped (ExcelJS floatifies cfvo val)', async () => {
  const out = await roundTrip([
    {
      type: 'colorScale',
      priority: 1,
      cfvo: [{ type: 'min' }, { type: 'formula', value: '=A1*2' }, { type: 'max' }],
      color: [{ argb: 'FFF8696B' }, { argb: 'FFFFEB84' }, { argb: 'FF63BE7B' }],
    },
  ]);
  // The formula stop can't round-trip, so the whole rule is skipped rather than
  // emitted with a corrupt "NaN" threshold.
  assert.equal(out.length, 0);
});

test('a foreign rule with no style does not throw the export', async () => {
  // A resource payload from another tool / future version may omit `style`.
  // Export must tolerate it (cfStyleToDxf guards null) rather than aborting save.
  const snapshot = {
    resources: [
      {
        name: CONDITIONAL_FORMATTING_RESOURCE,
        data: JSON.stringify({
          'sheet-1': [
            {
              cfId: 'x',
              ranges: [{ startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }],
              stopIfTrue: false,
              rule: { type: 'highlightCell', subType: 'number', operator: 'greaterThan', value: 5 },
            },
          ],
        }),
      },
    ],
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
  } as any;
  const rules = readConditionalFormattingFromSnapshot(snapshot)['sheet-1'] ?? [];
  assert.equal(rules.length, 1);
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('S');
  assert.doesNotThrow(() => applyConditionalFormattingToXlsxWorksheet(ws, rules));
});
