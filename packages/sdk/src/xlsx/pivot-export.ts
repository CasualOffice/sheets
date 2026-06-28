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

/**
 * Generate native `xl/pivotTables` + `xl/pivotCaches` OOXML for an
 * in-app-created pivot, so Excel re-opens it as a real PivotTable (not flat
 * cells). The output is a `PivotPassthroughPayload` — the SAME shape the
 * import-side passthrough produces — so the existing `applyPivotsToZip`
 * injects it (rel renumbering, Content-Types, workbook.xml `<pivotCaches>`,
 * sheet rels). We only build the *parts*, not the injection layer.
 *
 * Scope (A.1 spike): one row field + one value field, single worksheet
 * source, no column/filter/grouping. `refreshOnLoad="1"` is set so Excel
 * rebuilds the cache from the live source on open — the cache records we
 * emit only need to be structurally valid, not pixel-perfect.
 *
 * Pure + framework-free so the XML is exhaustively unit-testable; a fixture
 * script assembles a real .xlsx for the Excel-open validation this needs.
 */

import type { PivotPassthroughPayload } from './pivot-passthrough';

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL_TYPE_CACHE_DEF =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition';
const REL_TYPE_CACHE_REC =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords';

export type PivotAgg = 'sum' | 'count' | 'average' | 'min' | 'max' | 'distinctCount';

export interface NativePivotInput {
  /** Source + output worksheet name (single-sheet spike). */
  sheetName: string;
  /** Source range in A1 form, header row included (e.g. `A1:C5`). */
  sourceRef: string;
  /** Field names from the source header row. */
  headers: string[];
  /** Data rows (header excluded), row-major, aligned to `headers`. */
  records: Array<Array<string | number | null>>;
  /** Source column index used as the (single) row field. */
  rowField: number;
  /** Source column index aggregated as the (single) value field. */
  valueField: number;
  /** Aggregation for the value field. */
  valueAgg: PivotAgg;
  /** Top-left of the pivot output, A1 form (e.g. `E1`). */
  targetRef: string;
  /** Numbering — lets multiple pivots / imported caches coexist. */
  cacheId?: number;
  index?: number;
}

const esc = (s: string): string =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

const isNumeric = (v: unknown): v is number => typeof v === 'number' && Number.isFinite(v);

/** Subtotal attribute Excel expects on `<dataField>` (sum is the default → omit). */
const SUBTOTAL: Record<PivotAgg, string> = {
  sum: '',
  count: ' subtotal="count"',
  average: ' subtotal="average"',
  min: ' subtotal="min"',
  max: ' subtotal="max"',
  distinctCount: ' subtotal="countNums"',
};

const AGG_LABEL: Record<PivotAgg, string> = {
  sum: 'Sum',
  count: 'Count',
  average: 'Average',
  min: 'Min',
  max: 'Max',
  distinctCount: 'Distinct Count',
};

/** Column letter for a 0-based index — A1 helpers for the location ref. */
function colLetters(col: number): string {
  let n = col + 1;
  let s = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function parseCellRef(a1: string): { row: number; col: number } {
  const m = /^([A-Za-z]+)(\d+)$/.exec(a1.trim());
  if (!m) return { row: 0, col: 0 };
  let col = 0;
  for (const ch of m[1].toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { row: Number(m[2]) - 1, col: col - 1 };
}

/** UTF-8 → base64, portable across the worker (TextEncoder/btoa) and Node. */
function toBase64(str: string): string {
  if (typeof Buffer !== 'undefined') return Buffer.from(str, 'utf8').toString('base64');
  const bytes = new TextEncoder().encode(str);
  let bin = '';
  for (const b of bytes) bin += String.fromCharCode(b);
  return btoa(bin);
}

const xmlDecl = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const relsOpen = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`;

/**
 * Build the OOXML parts for a single-row-field, single-value pivot, returned
 * as a `PivotPassthroughPayload` ready for `applyPivotsToZip`.
 */
export function generateNativePivot(input: NativePivotInput): PivotPassthroughPayload {
  const n = input.index ?? 1;
  const cacheId = input.cacheId ?? n;
  const colCount = input.headers.length;

  // Classify each source column: string columns get shared items (records
  // reference them by index); numeric columns inline their values.
  const isNumCol = input.headers.map(
    (_, c) =>
      input.records.length > 0 && input.records.every((r) => r[c] == null || isNumeric(r[c])),
  );
  // Distinct shared items per string column, in first-appearance order.
  const shared: string[][] = input.headers.map((_, c) => {
    if (isNumCol[c]) return [];
    const seen = new Set<string>();
    const out: string[] = [];
    for (const r of input.records) {
      const v = r[c] == null ? '' : String(r[c]);
      if (!seen.has(v)) {
        seen.add(v);
        out.push(v);
      }
    }
    return out;
  });

  // ── pivotCacheDefinition{n}.xml ──────────────────────────────────────────
  const cacheFields = input.headers
    .map((h, c) => {
      if (isNumCol[c]) {
        const nums = input.records.map((r) => r[c]).filter(isNumeric);
        const min = nums.length ? Math.min(...nums) : 0;
        const max = nums.length ? Math.max(...nums) : 0;
        return (
          `<cacheField name="${esc(h)}" numFmtId="0">` +
          `<sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" minValue="${min}" maxValue="${max}"/>` +
          `</cacheField>`
        );
      }
      const items = shared[c].map((v) => `<s v="${esc(v)}"/>`).join('');
      return (
        `<cacheField name="${esc(h)}" numFmtId="0">` +
        `<sharedItems count="${shared[c].length}">${items}</sharedItems>` +
        `</cacheField>`
      );
    })
    .join('');

  const cacheDef =
    `${xmlDecl}\n` +
    `<pivotCacheDefinition xmlns="${NS_MAIN}" xmlns:r="${NS_R}" r:id="rId1" refreshOnLoad="1" refreshedBy="Casual Sheets" createdVersion="6" refreshedVersion="6" minRefreshableVersion="3" recordCount="${input.records.length}">` +
    `<cacheSource type="worksheet"><worksheetSource ref="${esc(input.sourceRef)}" sheet="${esc(input.sheetName)}"/></cacheSource>` +
    `<cacheFields count="${colCount}">${cacheFields}</cacheFields>` +
    `</pivotCacheDefinition>`;

  // ── pivotCacheRecords{n}.xml ─────────────────────────────────────────────
  const recXml = input.records
    .map((r) => {
      const cells = r
        .map((v, c) => {
          if (isNumCol[c]) return isNumeric(v) ? `<n v="${v}"/>` : `<m/>`;
          const idx = shared[c].indexOf(v == null ? '' : String(v));
          return `<x v="${idx < 0 ? 0 : idx}"/>`;
        })
        .join('');
      return `<r>${cells}</r>`;
    })
    .join('');
  const cacheRec =
    `${xmlDecl}\n` +
    `<pivotCacheRecords xmlns="${NS_MAIN}" xmlns:r="${NS_R}" count="${input.records.length}">${recXml}</pivotCacheRecords>`;

  // ── pivotTable{n}.xml ────────────────────────────────────────────────────
  const rowItemCount = shared[input.rowField].length; // distinct row values
  // location: header row + one row per distinct value + grand total, 2 cols.
  const tgt = parseCellRef(input.targetRef);
  const endRow = tgt.row + rowItemCount + 1; // +header +grandTotal − 1 ... see below
  const ref = `${colLetters(tgt.col)}${tgt.row + 1}:${colLetters(tgt.col + 1)}${endRow + 1}`;

  const pivotFields = input.headers
    .map((_, c) => {
      if (c === input.rowField) {
        const items = shared[c].map((_, i) => `<item x="${i}"/>`).join('') + `<item t="default"/>`;
        return `<pivotField axis="axisRow" showAll="0"><items count="${shared[c].length + 1}">${items}</items></pivotField>`;
      }
      if (c === input.valueField) return `<pivotField dataField="1" showAll="0"/>`;
      return `<pivotField showAll="0"/>`;
    })
    .join('');

  const rowItems =
    shared[input.rowField]
      .map((_, i) => (i === 0 ? `<i><x/></i>` : `<i><x v="${i}"/></i>`))
      .join('') + `<i t="grand"><x/></i>`;

  const dataName = `${AGG_LABEL[input.valueAgg]} of ${input.headers[input.valueField] ?? 'Value'}`;
  const pivotTable =
    `${xmlDecl}\n` +
    `<pivotTableDefinition xmlns="${NS_MAIN}" xmlns:r="${NS_R}" name="PivotTable${n}" cacheId="${cacheId}" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" updatedVersion="6" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="6" indent="0" outline="1" outlineData="1" multipleFieldFilters="0">` +
    `<location ref="${ref}" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/>` +
    `<pivotFields count="${colCount}">${pivotFields}</pivotFields>` +
    `<rowFields count="1"><field x="${input.rowField}"/></rowFields>` +
    `<rowItems count="${rowItemCount + 1}">${rowItems}</rowItems>` +
    `<colItems count="1"><i/></colItems>` +
    `<dataFields count="1"><dataField name="${esc(dataName)}" fld="${input.valueField}" baseField="0" baseItem="0"${SUBTOTAL[input.valueAgg]}/></dataFields>` +
    `<pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/>` +
    `</pivotTableDefinition>`;

  // ── per-part .rels ───────────────────────────────────────────────────────
  const cacheDefRels =
    `${xmlDecl}\n${relsOpen}` +
    `<Relationship Id="rId1" Type="${REL_TYPE_CACHE_REC}" Target="pivotCacheRecords${n}.xml"/>` +
    `</Relationships>`;
  const pivotTableRels =
    `${xmlDecl}\n${relsOpen}` +
    `<Relationship Id="rId1" Type="${REL_TYPE_CACHE_DEF}" Target="../pivotCaches/pivotCacheDefinition${n}.xml"/>` +
    `</Relationships>`;

  const b64 = toBase64;

  return {
    parts: {
      [`xl/pivotCaches/pivotCacheDefinition${n}.xml`]: b64(cacheDef),
      [`xl/pivotCaches/pivotCacheRecords${n}.xml`]: b64(cacheRec),
      [`xl/pivotCaches/_rels/pivotCacheDefinition${n}.xml.rels`]: b64(cacheDefRels),
      [`xl/pivotTables/pivotTable${n}.xml`]: b64(pivotTable),
      [`xl/pivotTables/_rels/pivotTable${n}.xml.rels`]: b64(pivotTableRels),
    },
    workbookPivotCachesXml: `<pivotCaches><pivotCache cacheId="${cacheId}" r:id="rIdCache${n}"/></pivotCaches>`,
    workbookCacheRels: [
      { origId: `rIdCache${n}`, target: `pivotCaches/pivotCacheDefinition${n}.xml` },
    ],
    perSheet: [
      {
        sheetName: input.sheetName,
        pivotTableRels: [{ origId: `rIdTbl${n}`, target: `../pivotTables/pivotTable${n}.xml` }],
      },
    ],
  };
}
