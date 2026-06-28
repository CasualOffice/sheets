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
import * as fs from 'node:fs';
import * as path from 'node:path';
import { createRequire } from 'node:module';
import { waitForUniver } from './_helpers';

const require = createRequire(import.meta.url);
type OdsModule = typeof import('../../apps/web/src/ods');
type OdsWorkbookData = Parameters<OdsModule['workbookDataToOds']>[0];
// `@e965/xlsx` is installed in the `apps/web` workspace, not the repo root.
// Resolve from there so Playwright can run this spec from the monorepo root.
// eslint-disable-next-line @typescript-eslint/no-var-requires
const XLSX = require(require.resolve('@e965/xlsx', { paths: [path.join(process.cwd(), 'apps/web')] }));

declare global {
  interface Window {
    __odsAudit?: {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      odsToWorkbookData: (buf: ArrayBuffer) => Promise<any>;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      workbookDataToOds: (data: any) => Promise<Blob>;
    };
  }
}

async function exposeConverters(page: Page) {
  await page.evaluate(async () => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const mod = await import(/* @vite-ignore */ '/src/ods/index.ts' as any);
    window.__odsAudit = mod;
  });
}

function buildReferenceOds(): Buffer {
  const data: XLSX.WorkSheet = {
    A1: { t: 's', v: 'Quarter' },
    B1: { t: 's', v: 'Revenue' },
    C1: { t: 's', v: 'Profile' },
    D1: { t: 's', v: 'Notes' },
    A2: { t: 's', v: 'Q1' },
    B2: { t: 'n', v: 1234.56, z: '"$"#,##0.00' },
    C2: { t: 's', v: 'GitHub', l: { Target: 'https://github.com/CasualOffice/sheets' } },
    D2: { t: 's', v: 'memo', c: [{ a: 'audit', t: 'ODS comment' }] },
    A3: { t: 's', v: 'Q2' },
    B3: { t: 'n', v: 0.25, z: '0.00%' },
    A4: { t: 's', v: 'Total' },
    B4: { t: 'n', v: 1234.81, f: 'SUM(B2:B3)' },
    C4: { t: 's', v: 'Cross' },
    D4: { t: 's', v: 'shadow', f: 'Hidden!A1' },
    A6: { t: 's', v: 'Merged banner' },
    '!merges': [XLSX.utils.decode_range('A6:B6')],
    '!cols': [{ wpx: 125, wch: 20 }, { wpx: 77, wch: 12 }],
    '!rows': [{ hpx: 24, hpt: 18 }, { hpx: 18, hpt: 13.5 }],
    '!ref': 'A1:D6',
  };

  const hidden: XLSX.WorkSheet = {
    A1: { t: 's', v: 'shadow' },
    '!ref': 'A1:A1',
  };

  const wb = XLSX.utils.book_new();
  wb.Workbook = {
    Names: [{ Name: 'RevenueCell', Ref: 'Data!$B$2' }],
  };
  XLSX.utils.book_append_sheet(wb, data, 'Data');
  XLSX.utils.book_append_sheet(wb, hidden, 'Hidden');
  return Buffer.from(XLSX.write(wb, { type: 'buffer', bookType: 'ods' }));
}

type Probe = {
  category: string;
  what: string;
  reference: unknown;
  actual: unknown;
  result: boolean | 'partial';
};

type StyleRoundTripResult = {
  bl?: number;
  cl?: { rgb?: string };
  bg?: { rgb?: string };
  ht?: number;
  vt?: number;
  n?: { pattern?: string };
} | null;

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

function readOdsWorkbook(bytes: Buffer): XLSX.WorkBook {
  return XLSX.read(bytes, {
    type: 'buffer',
    cellFormula: true,
    cellNF: true,
    cellHTML: true,
    cellText: true,
  });
}

function compareWorkbooks(ref: XLSX.WorkBook, got: XLSX.WorkBook, gotSnapshot?: {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  sheetOrder?: string[];
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  sheets?: Record<string, any>;
}): Probe[] {
  const probes: Probe[] = [];
  const push = (
    category: string,
    what: string,
    reference: unknown,
    actual: unknown,
    matcher: (r: unknown, a: unknown) => boolean | 'partial' = (r, a) =>
      Object.is(r, a) || JSON.stringify(r) === JSON.stringify(a),
  ) => probes.push({ category, what, reference, actual, result: matcher(reference, actual) });

  push('Sheets', 'sheet order + names', ref.SheetNames, got.SheetNames);

  const refData = ref.Sheets.Data;
  const gotData = got.Sheets.Data;
  if (!refData || !gotData) {
    push('Sheets', 'Data sheet present', true, Boolean(gotData));
    return probes;
  }

  push('Values', 'A1 header', refData.A1?.v, gotData.A1?.v);
  push('Values', 'B2 number', refData.B2?.v, gotData.B2?.v);
  push('Formulas', 'B4 SUM formula text', refData.B4?.f, gotData.B4?.f);
  push('Formulas', 'B4 cached result', refData.B4?.v, gotData.B4?.v);
  push('Formulas', 'D4 cross-sheet formula text', refData.D4?.f, gotData.D4?.f);
  push('Formulas', 'D4 cross-sheet cached result', refData.D4?.v, gotData.D4?.v);
  push(
    'Merges',
    'A6:B6 merged',
    true,
    Boolean(gotData['!merges']?.some((m) => m.s.r === 5 && m.s.c === 0 && m.e.r === 5 && m.e.c === 1)),
  );
  push('Number format', 'B2 currency', refData.B2?.z, gotData.B2?.z);
  push('Number format', 'B3 percent', refData.B3?.z, gotData.B3?.z);
  push(
    'Hyperlinks',
    'C2 external hyperlink',
    refData.C2?.l?.Target,
    gotData.C2?.l?.Target,
  );
  push(
    'Comments',
    'D2 comment text',
    refData.D2?.c?.map((c) => c.t).join('\n'),
    gotData.D2?.c?.map((c) => c.t).join('\n'),
  );
  push(
    'Defined names',
    'RevenueCell',
    ref.Workbook?.Names?.find((n) => n.Name === 'RevenueCell')?.Ref ?? null,
    got.Workbook?.Names?.find((n) => n.Name === 'RevenueCell')?.Ref ?? null,
  );
  const gotSheetId = gotSnapshot?.sheetOrder?.[0];
  const gotSheet = gotSheetId ? gotSnapshot?.sheets?.[gotSheetId] : undefined;
  push('Dimensions', 'column A width', 125, gotSheet?.columnData?.[0]?.w ?? null);
  push('Dimensions', 'column B width', 77, gotSheet?.columnData?.[1]?.w ?? null);
  push('Dimensions', 'row 1 height', 24, gotSheet?.rowData?.[0]?.h ?? null);
  push('Dimensions', 'row 2 height', 18, gotSheet?.rowData?.[1]?.h ?? null);

  return probes;
}

function appendStyleProbes(probes: Probe[], style: StyleRoundTripResult): void {
  const push = (
    category: string,
    what: string,
    reference: unknown,
    actual: unknown,
    matcher: (r: unknown, a: unknown) => boolean | 'partial' = (r, a) =>
      Object.is(r, a) || JSON.stringify(r) === JSON.stringify(a),
  ) => probes.push({ category, what, reference, actual, result: matcher(reference, actual) });

  push('Styles · font', 'bold preserved', 1, style?.bl ?? null);
  push('Styles · font', 'font color preserved', '#ff0000', style?.cl?.rgb ?? null);
  push('Styles · fill', 'fill color preserved', '#00ff00', style?.bg?.rgb ?? null);
  push('Styles · alignment', 'horizontal alignment preserved', 2, style?.ht ?? null);
  push('Styles · alignment', 'vertical alignment preserved', 2, style?.vt ?? null);
}

function buildReport(probes: Probe[]): string {
  const byCat = new Map<string, Probe[]>();
  for (const p of probes) {
    const arr = byCat.get(p.category) ?? [];
    arr.push(p);
    byCat.set(p.category, arr);
  }

  const totals = { ok: 0, partial: 0, miss: 0 };
  for (const p of probes) {
    if (p.result === true) totals.ok++;
    else if (p.result === 'partial') totals.partial++;
    else totals.miss++;
  }

  const lines: string[] = [];
  lines.push('# ODS round-trip lossiness audit');
  lines.push('');
  lines.push('Generated by `tests/e2e/ods-lossiness-audit.spec.ts`. Build a reference ODS');
  lines.push('with SheetJS, run it through our parser → snapshot → exporter pipeline,');
  lines.push('re-parse the result with SheetJS, and compare feature-by-feature.');
  lines.push('');
  lines.push('Legend: `✅` survived · `⚠️` partial · `❌` dropped');
  lines.push('');
  lines.push(`**Totals**: ${totals.ok} ✅ · ${totals.partial} ⚠️ · ${totals.miss} ❌ (of ${probes.length} probes)`);
  lines.push('');
  lines.push('## Parser constraints');
  lines.push('');
  lines.push('- Current `@e965/xlsx` ODS parser exposes formulas, merges, hyperlinks, comments, defined names, and number formats.');
  lines.push('- Current probing does not show ODS row heights, column widths, freeze panes, or general font/fill/alignment style objects on the parsed worksheet model.');
  lines.push('- Those features need a separate viability pass before we can claim true round-trip support in our app.');
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

test.describe('ods round-trip lossiness audit', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await exposeConverters(page);
  });

  test('build report + assert current supported ODS features have not regressed', async ({ page }) => {
    const referenceBytes = buildReferenceOds();
    const referenceWb = readOdsWorkbook(referenceBytes);

    const result = await page.evaluate(async (bytesArr) => {
      const buf = new Uint8Array(bytesArr).buffer;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const snapshot: any = await window.__odsAudit!.odsToWorkbookData(buf);
      const blob = await window.__odsAudit!.workbookDataToOds(snapshot);
      const out = new Uint8Array(await blob.arrayBuffer());
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const reloaded: any = await window.__odsAudit!.odsToWorkbookData(out.buffer);
      return { bytes: Array.from(out), snapshot: reloaded };
    }, Array.from(referenceBytes));

    const gotWb = readOdsWorkbook(Buffer.from(result.bytes));
    const probes = compareWorkbooks(referenceWb, gotWb, result.snapshot);
    const styleRoundTrip = await page.evaluate(async () => {
      const snapshot: OdsWorkbookData = {
        id: 'audit-style-1',
        rev: 1,
        name: 'wb',
        appVersion: '0.22.1',
        locale: 1,
        styles: {
          s0: {
            bl: 1,
            cl: { rgb: '#ff0000' },
            bg: { rgb: '#00ff00' },
            ht: 2,
            vt: 2,
            n: { pattern: '0.00%' },
          },
        },
        sheetOrder: ['s1'],
        sheets: {
          s1: {
            id: 's1',
            name: 'Data',
            cellData: {
              0: {
                0: { v: 0.25, s: 's0' },
              },
            },
            rowCount: 100,
            columnCount: 26,
          },
        },
      };
      const blob = await window.__odsAudit!.workbookDataToOds(snapshot);
      const reloaded = await window.__odsAudit!.odsToWorkbookData(await blob.arrayBuffer());
      const sheetId = reloaded.sheetOrder[0];
      const cell = reloaded.sheets[sheetId]?.cellData?.[0]?.[0];
      return (cell?.s ? reloaded.styles?.[cell.s] : null) as StyleRoundTripResult;
    });
    appendStyleProbes(probes, styleRoundTrip);
    const report = buildReport(probes);

    const outPath = path.join(process.cwd(), 'docs', 'ods-lossiness.md');
    fs.mkdirSync(path.dirname(outPath), { recursive: true });
    fs.writeFileSync(outPath, report, 'utf8');

    console.info('\n' + report);

    const lock = (category: string, what: string) => {
      const p = probes.find((x) => x.category === category && x.what === what);
      expect(p, `${category} / ${what} probe missing`).toBeDefined();
      expect(p!.result, `${category} / ${what} regressed`).toBe(true);
    };

    lock('Sheets', 'sheet order + names');
    lock('Values', 'A1 header');
    lock('Values', 'B2 number');
    lock('Formulas', 'B4 SUM formula text');
    lock('Formulas', 'B4 cached result');
    lock('Formulas', 'D4 cross-sheet formula text');
    lock('Formulas', 'D4 cross-sheet cached result');
    lock('Merges', 'A6:B6 merged');
    lock('Number format', 'B2 currency');
    lock('Number format', 'B3 percent');
    lock('Hyperlinks', 'C2 external hyperlink');
    lock('Dimensions', 'column A width');
    lock('Dimensions', 'column B width');
    lock('Dimensions', 'row 1 height');
    lock('Dimensions', 'row 2 height');
    lock('Styles · font', 'bold preserved');
    lock('Styles · font', 'font color preserved');
    lock('Styles · fill', 'fill color preserved');
    lock('Styles · alignment', 'horizontal alignment preserved');
    lock('Styles · alignment', 'vertical alignment preserved');
    lock('Defined names', 'RevenueCell');
  });
});
