import { expect, test, type Page } from '@playwright/test';
import ExcelJS from 'exceljs';
import * as fs from 'node:fs';
import * as path from 'node:path';
import { waitForUniver } from './_helpers';

/**
 * xlsx round-trip lossiness audit.
 *
 * Build a rich xlsx in Node with ExcelJS, run it through our
 *
 *     parser  →  IWorkbookData snapshot  →  exporter
 *
 * pipeline, then re-parse the result in Node with ExcelJS and diff
 * the two workbooks feature-by-feature. The output is a single
 * markdown report at `docs/xlsx-lossiness.md` that lists what we
 * keep, what we drop, and what we partially mangle — driving the
 * fix punch list before we open the Univer fork.
 *
 * This spec is the audit harness. The categories below are an
 * incremental list — adding a new one is one new function +
 * one new section in the report.
 */

declare global {
  interface Window {
    __xlsx?: {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      xlsxToWorkbookData: (buf: ArrayBuffer) => Promise<any>;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      workbookDataToXlsx: (data: any) => Promise<Blob>;
    };
  }
}

async function exposeConverters(page: Page) {
  await page.evaluate(async () => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const mod = await import(/* @vite-ignore */ '/src/xlsx/index.ts' as any);
    window.__xlsx = mod;
  });
}

/** Build the "kitchen sink" reference workbook. Everything we want to
 *  measure goes in here — adding a new probe is one new write below. */
async function buildReferenceXlsx(): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'casual-sheets audit';
  wb.created = new Date('2026-01-15T00:00:00Z');

  // Sheet 1 — values, formulas, styles, merges, hyperlinks, comments.
  const s1 = wb.addWorksheet('Data', { properties: { tabColor: { argb: 'FFFF5722' } } });
  s1.views = [{ state: 'frozen', xSplit: 1, ySplit: 1, activeCell: 'B2' }];
  s1.getColumn(1).width = 18;
  s1.getColumn(2).width = 12;
  s1.getRow(1).height = 24;

  s1.getCell('A1').value = 'Quarter';
  s1.getCell('B1').value = 'Revenue';
  s1.getCell('C1').value = 'Notes';
  s1.getCell('A2').value = 'Q1';
  s1.getCell('B2').value = 1234.56;
  s1.getCell('A3').value = 'Q2';
  s1.getCell('B3').value = 2345.67;
  s1.getCell('A4').value = 'Q3';
  s1.getCell('B4').value = 3456.78;
  s1.getCell('A5').value = 'Total';
  s1.getCell('B5').value = { formula: 'SUM(B2:B4)', result: 7037.01 };

  // Cross-sheet formula (filled in after sheet 2 exists).

  // Cell styling on the header row.
  ['A1', 'B1', 'C1'].forEach((addr) => {
    const c = s1.getCell(addr);
    c.font = { name: 'Calibri', size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2563EB' } };
    c.alignment = { horizontal: 'center', vertical: 'middle' };
    c.border = {
      top: { style: 'thin', color: { argb: 'FF1F2937' } },
      bottom: { style: 'thin', color: { argb: 'FF1F2937' } },
      left: { style: 'thin', color: { argb: 'FF1F2937' } },
      right: { style: 'thin', color: { argb: 'FF1F2937' } },
    };
  });

  // Number formats — currency, percent, date.
  s1.getCell('B2').numFmt = '"$"#,##0.00';
  s1.getCell('B3').numFmt = '0.00%';
  s1.getCell('B4').numFmt = 'yyyy-mm-dd';

  // Merge.
  s1.mergeCells('A7:C7');
  s1.getCell('A7').value = 'Merged Banner';
  s1.getCell('A7').alignment = { horizontal: 'center' };

  // Hyperlink.
  s1.getCell('C2').value = { text: 'GitHub', hyperlink: 'https://github.com/schnsrw/sheets' };

  // Comment on B2 (ExcelJS API).
  s1.getCell('B2').note = 'This is a comment on Q1 revenue.';

  // Sheet 2 — second sheet for cross-sheet formula + hidden state.
  const s2 = wb.addWorksheet('Hidden');
  s2.getCell('A1').value = 'concealed';
  s2.state = 'hidden';

  // Cross-sheet formula references the hidden sheet.
  s1.getCell('C5').value = { formula: "Hidden!A1", result: 'concealed' };

  // Defined name (named range).
  wb.definedNames.add("Data!$B$2:$B$4", 'RevenueRange');

  // Data validation — list constraint.
  s1.dataValidations.add('D2:D4', {
    type: 'list',
    allowBlank: true,
    formulae: ['"alpha,beta,gamma"'],
  });

  // Page setup — margins + orientation.
  s1.pageSetup = {
    orientation: 'landscape',
    paperSize: 9,
    margins: { left: 0.75, right: 0.75, top: 1, bottom: 1, header: 0.3, footer: 0.3 },
  };
  s1.headerFooter = { oddHeader: '&LCasual Sheets Audit', oddFooter: '&CPage &P of &N' };

  // Sheet 3 — table + autofilter (cleanly different scope).
  const s3 = wb.addWorksheet('Tabular');
  s3.addTable({
    name: 'PeopleTable',
    ref: 'A1',
    headerRow: true,
    style: { theme: 'TableStyleMedium2', showRowStripes: true },
    columns: [
      { name: 'Name', filterButton: true },
      { name: 'Age', filterButton: true },
    ],
    rows: [
      ['Alice', 30],
      ['Bob', 25],
      ['Carol', 41],
    ],
  });

  const arr = await wb.xlsx.writeBuffer();
  return Buffer.from(arr);
}

type Probe = {
  category: string;
  what: string;
  reference: unknown;
  actual: unknown;
  /** Did the round-trip preserve this? `true` = full, `false` = lost,
   *  `'partial'` = present but different shape (still flagged). */
  result: boolean | 'partial';
};

function fmtVal(v: unknown): string {
  if (v === null || v === undefined) return '∅';
  if (typeof v === 'string') return JSON.stringify(v);
  if (typeof v === 'object') {
    try {
      return JSON.stringify(v);
    } catch {
      return String(v);
    }
  }
  return String(v);
}

function compareWorkbooks(ref: ExcelJS.Workbook, got: ExcelJS.Workbook): Probe[] {
  const probes: Probe[] = [];
  const push = (
    category: string,
    what: string,
    reference: unknown,
    actual: unknown,
    matcher: (r: unknown, a: unknown) => boolean | 'partial' = (r, a) => Object.is(r, a) || JSON.stringify(r) === JSON.stringify(a),
  ) => probes.push({ category, what, reference, actual, result: matcher(reference, actual) });

  // Sheet identity.
  push(
    'Sheets',
    'sheet order + names',
    ref.worksheets.map((s) => s.name),
    got.worksheets.map((s) => s.name),
  );

  const refData = ref.getWorksheet('Data');
  const gotData = got.getWorksheet('Data');
  if (!refData || !gotData) {
    push('Sheets', 'Data sheet present', true, Boolean(gotData));
    return probes;
  }

  // Values.
  push('Values', 'A1 header', refData.getCell('A1').value, gotData.getCell('A1').value);
  push('Values', 'B2 number', refData.getCell('B2').value, gotData.getCell('B2').value);
  push(
    'Formulas',
    'B5 SUM formula text',
    (refData.getCell('B5').value as { formula?: string })?.formula,
    (gotData.getCell('B5').value as { formula?: string } | string)
      ? typeof gotData.getCell('B5').value === 'object'
        ? (gotData.getCell('B5').value as { formula?: string }).formula
        : String(gotData.getCell('B5').value).replace(/^=/, '')
      : null,
  );
  push(
    'Formulas',
    'C5 cross-sheet formula references Hidden!A1',
    'Hidden!A1',
    (() => {
      const v = gotData.getCell('C5').value;
      if (typeof v === 'object' && v && 'formula' in v) return (v as { formula?: string }).formula;
      if (typeof v === 'string') return v.replace(/^=/, '');
      return null;
    })(),
  );

  // Styles on the header.
  const refHdr = refData.getCell('A1');
  const gotHdr = gotData.getCell('A1');
  push('Styles · font', 'bold preserved', refHdr.font?.bold, gotHdr.font?.bold);
  push('Styles · font', 'size preserved', refHdr.font?.size, gotHdr.font?.size);
  push('Styles · font', 'family preserved', refHdr.font?.name, gotHdr.font?.name);
  push(
    'Styles · font',
    'color preserved',
    refHdr.font?.color,
    gotHdr.font?.color,
    (r, a) => {
      // ExcelJS sometimes returns theme indices vs argb depending on the writer.
      // We treat "any non-default color present" as partial when shape differs.
      if (JSON.stringify(r) === JSON.stringify(a)) return true;
      if (r && a) return 'partial';
      return false;
    },
  );
  push(
    'Styles · fill',
    'header fill preserved',
    refHdr.fill,
    gotHdr.fill,
    (r, a) => {
      if (!r || !a) return Boolean(r) === Boolean(a);
      if (JSON.stringify(r) === JSON.stringify(a)) return true;
      // Pattern/color may serialize slightly differently; flag as partial if both are solid fills.
      const rt = (r as { type?: string }).type;
      const at = (a as { type?: string }).type;
      if (rt === at) return 'partial';
      return false;
    },
  );
  push('Styles · alignment', 'header horizontal', refHdr.alignment?.horizontal, gotHdr.alignment?.horizontal);
  push('Styles · alignment', 'header vertical', refHdr.alignment?.vertical, gotHdr.alignment?.vertical);
  push(
    'Styles · border',
    'header borders (4 sides) preserved',
    Object.keys(refHdr.border ?? {}).filter((k) => Boolean((refHdr.border as Record<string, unknown>)[k])).sort(),
    Object.keys(gotHdr.border ?? {}).filter((k) => Boolean((gotHdr.border as Record<string, unknown>)[k])).sort(),
  );

  // Number formats.
  push('Number format', 'B2 currency', refData.getCell('B2').numFmt, gotData.getCell('B2').numFmt);
  push('Number format', 'B3 percent', refData.getCell('B3').numFmt, gotData.getCell('B3').numFmt);
  push('Number format', 'B4 date', refData.getCell('B4').numFmt, gotData.getCell('B4').numFmt);

  // Merge.
  push(
    'Merges',
    'A7:C7 merged',
    true,
    (() => {
      // Walk merge ranges — ExcelJS exposes via `_merges` private; use mergeCells reverse lookup.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const merges = (gotData as any).model?.merges as string[] | undefined;
      return Boolean(merges?.some((m) => m.includes('A7')));
    })(),
  );

  // Hyperlink.
  push(
    'Hyperlinks',
    'C2 → github.com/schnsrw/sheets',
    'https://github.com/schnsrw/sheets',
    (() => {
      const v = gotData.getCell('C2').value;
      if (typeof v === 'object' && v && 'hyperlink' in v) return (v as { hyperlink?: string }).hyperlink;
      return null;
    })(),
  );

  // Comment.
  push(
    'Comments',
    'B2 comment text',
    'This is a comment on Q1 revenue.',
    (() => {
      const note = gotData.getCell('B2').note;
      if (typeof note === 'string') return note;
      if (note && typeof note === 'object' && 'texts' in note) {
        return (note as { texts?: { text?: string }[] }).texts?.map((t) => t.text ?? '').join('') ?? null;
      }
      return null;
    })(),
  );

  // Defined name.
  push(
    'Defined names',
    'RevenueRange',
    'Data!$B$2:$B$4',
    (() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const refs = (got.definedNames as any).getRanges?.('RevenueRange') as { ranges?: string[] } | undefined;
      return refs?.ranges?.[0] ?? null;
    })(),
  );

  // Data validation.
  push(
    'Data validation',
    'D2:D4 list constraint',
    'list',
    (() => {
      const dv = gotData.dataValidations;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const map = (dv as any).model as Record<string, { type?: string }> | undefined;
      if (!map) return null;
      const key = Object.keys(map).find((k) => k.startsWith('D2') || k.includes('D2:D4'));
      return key ? map[key]?.type ?? null : null;
    })(),
  );

  // Sheet props on Data.
  push('Sheet props', 'tab color (Data)', { argb: 'FFFF5722' }, refData.properties.tabColor);
  push('Sheet props', 'tab color survives round-trip', { argb: 'FFFF5722' }, gotData.properties.tabColor);
  push(
    'Sheet props',
    'frozen rows',
    1,
    (gotData.views?.[0] as { ySplit?: number } | undefined)?.ySplit,
  );
  push(
    'Sheet props',
    'frozen columns',
    1,
    (gotData.views?.[0] as { xSplit?: number } | undefined)?.xSplit,
  );

  // Hidden sheet.
  const gotHidden = got.getWorksheet('Hidden');
  push('Sheet props', 'hidden sheet survives', 'hidden', gotHidden?.state);

  // Column widths + row heights.
  push('Dimensions', 'column A width (18)', 18, gotData.getColumn(1).width);
  push('Dimensions', 'column B width (12)', 12, gotData.getColumn(2).width);
  push('Dimensions', 'row 1 height (24)', 24, gotData.getRow(1).height);

  // Page setup.
  push('Page setup', 'orientation landscape', 'landscape', gotData.pageSetup?.orientation);
  push(
    'Page setup',
    'header (left text)',
    '&LCasual Sheets Audit',
    gotData.headerFooter?.oddHeader,
  );

  // Tables.
  const gotTabular = got.getWorksheet('Tabular');
  push(
    'Tables',
    'PeopleTable defined',
    true,
    Boolean(gotTabular && gotTabular.getTables?.().length),
  );

  // Workbook metadata.
  push('Workbook metadata', 'creator', 'casual-sheets audit', got.creator);

  return probes;
}

function buildReport(probes: Probe[]): string {
  const byCat = new Map<string, Probe[]>();
  for (const p of probes) {
    const arr = byCat.get(p.category) ?? [];
    arr.push(p);
    byCat.set(p.category, arr);
  }
  const lines: string[] = [];
  lines.push('# xlsx round-trip lossiness audit');
  lines.push('');
  lines.push('Generated by `tests/e2e/xlsx-lossiness-audit.spec.ts`. Build a rich xlsx with');
  lines.push('ExcelJS, push it through our parser → snapshot → exporter pipeline, re-parse');
  lines.push('the result with ExcelJS, and compare feature-by-feature.');
  lines.push('');
  lines.push('Legend: `✅` survived · `⚠️` partial (present but shape differs) · `❌` dropped');
  lines.push('');

  const totals = { ok: 0, partial: 0, miss: 0 };
  for (const p of probes) {
    if (p.result === true) totals.ok++;
    else if (p.result === 'partial') totals.partial++;
    else totals.miss++;
  }
  lines.push(`**Totals**: ${totals.ok} ✅ · ${totals.partial} ⚠️ · ${totals.miss} ❌ (of ${probes.length} probes)`);
  lines.push('');

  for (const [cat, arr] of byCat) {
    lines.push(`## ${cat}`);
    lines.push('');
    lines.push('| Probe | Reference | Actual | Status |');
    lines.push('| --- | --- | --- | --- |');
    for (const p of arr) {
      const mark = p.result === true ? '✅' : p.result === 'partial' ? '⚠️' : '❌';
      lines.push(`| ${p.what} | \`${fmtVal(p.reference)}\` | \`${fmtVal(p.actual)}\` | ${mark} |`);
    }
    lines.push('');
  }

  return lines.join('\n');
}

test.describe('xlsx round-trip lossiness audit', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await exposeConverters(page);
  });

  test('build report + assert nothing already-fixed has regressed', async ({ page }) => {
    const referenceBytes = await buildReferenceXlsx();
    const referenceWb = new ExcelJS.Workbook();
    await referenceWb.xlsx.load(referenceBytes);

    // Push the reference through our parser → snapshot → exporter pipeline.
    const roundTripped = await page.evaluate(async (bytesArr) => {
      const buf = new Uint8Array(bytesArr).buffer;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const snapshot: any = await window.__xlsx!.xlsxToWorkbookData(buf);
      const blob = await window.__xlsx!.workbookDataToXlsx(snapshot);
      const out = new Uint8Array(await blob.arrayBuffer());
      return Array.from(out);
    }, Array.from(referenceBytes));

    const gotWb = new ExcelJS.Workbook();
    await gotWb.xlsx.load(Buffer.from(roundTripped));

    const probes = compareWorkbooks(referenceWb, gotWb);
    const report = buildReport(probes);

    // Always write the latest report so the doc reflects the build's actual state.
    const outPath = path.join(process.cwd(), 'docs', 'xlsx-lossiness.md');
    fs.mkdirSync(path.dirname(outPath), { recursive: true });
    fs.writeFileSync(outPath, report, 'utf8');

    // Surface the report in test output so CI logs it.
    console.info('\n' + report);

    // Locking gates — features that have explicit round-trip support today
    // (per existing specs). If any of these regress, fail loudly.
    const lock = (category: string, what: string) => {
      const p = probes.find((x) => x.category === category && x.what === what);
      expect(p, `${category} / ${what} probe missing`).toBeDefined();
      expect(p!.result, `${category} / ${what} regressed`).toBe(true);
    };
    lock('Values', 'A1 header');
    lock('Values', 'B2 number');
    lock('Sheets', 'sheet order + names');
    lock('Sheet props', 'frozen rows');
    lock('Sheet props', 'frozen columns');
    lock('Sheet props', 'hidden sheet survives');
    lock('Hyperlinks', 'C2 → github.com/schnsrw/sheets');
  });
});
