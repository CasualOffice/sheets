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
 * Post-process an exported .xlsx to add NATIVE pivot parts for in-app pivots,
 * so Excel re-opens them as real PivotTables. Runs after the SDK exporter,
 * loading the produced blob, generating `xl/pivotTables` / `xl/pivotCaches`
 * from each PivotModel + its source cells, and injecting via the same
 * `applyPivotsToZip` machinery the importer uses.
 *
 * Gated behind an OFF-by-default flag (`localStorage['cs-native-pivots']==='1'`
 * or `?nativePivots=1`) until the generated OOXML is confirmed to open cleanly
 * in real Excel (it can't be validated in CI). When off, export is unchanged —
 * in-app pivots still land as flat cells + the model resource. Scope (A.1):
 * single row field + single value field, no columns/filters/grouping; other
 * pivots are skipped (still exported as cells).
 */

import type { IWorkbookData } from '@univerjs/core';
import JSZip from 'jszip';
import { generateNativePivot, applyPivotsToZip } from '@casualoffice/sheets/xlsx';
import type { PivotModel } from '../pivots/types';

/** Opt-in flag — off by default. */
export function nativePivotsEnabled(): boolean {
  try {
    if (typeof localStorage !== 'undefined' && localStorage.getItem('cs-native-pivots') === '1') {
      return true;
    }
  } catch {
    /* storage blocked */
  }
  try {
    return new URLSearchParams(location.search).get('nativePivots') === '1';
  } catch {
    return false;
  }
}

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
const cellRef = (row: number, col: number) => `${colLetters(col)}${row + 1}`;

/** A pivot we can emit natively today: one row field, one value, no cols/filters. */
function isSimple(p: PivotModel): boolean {
  return (
    p.rows.length === 1 &&
    p.values.length === 1 &&
    (p.cols?.length ?? 0) === 0 &&
    (p.filters?.length ?? 0) === 0
  );
}

/** Read headers + records of a pivot's source range out of the snapshot. */
function readSource(
  data: IWorkbookData,
  p: PivotModel,
): { sheetName: string; headers: string[]; records: Array<Array<string | number | null>> } | null {
  const sheet = data.sheets?.[p.sourceSheetId];
  if (!sheet) return null;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const cd = (sheet.cellData ?? {}) as Record<number, Record<number, any>>;
  const { startRow, endRow, startColumn, endColumn } = p.source;
  const val = (r: number, c: number): string | number | null => {
    const v = cd[r]?.[c]?.v;
    return v == null ? null : (v as string | number);
  };
  const headers: string[] = [];
  for (let c = startColumn; c <= endColumn; c++) {
    const h = val(startRow, c);
    headers.push(h == null || h === '' ? `Column ${c - startColumn + 1}` : String(h));
  }
  const records: Array<Array<string | number | null>> = [];
  for (let r = startRow + 1; r <= endRow; r++) {
    const row: Array<string | number | null> = [];
    for (let c = startColumn; c <= endColumn; c++) row.push(val(r, c));
    records.push(row);
  }
  return { sheetName: sheet.name ?? '', headers, records };
}

/** Highest existing pivotTable index in the zip (imported pivots), or 0. */
function maxPivotIndex(zip: JSZip): number {
  let max = 0;
  for (const path of Object.keys(zip.files)) {
    const m = /pivotTable(\d+)\.xml$/.exec(path);
    if (m) max = Math.max(max, Number(m[1]));
  }
  return max;
}

/**
 * Inject native pivot parts for the simple in-app pivots into an exported blob.
 * Returns the original blob unchanged when the flag is off, nothing qualifies,
 * or anything throws (export must never fail because of this).
 */
export async function injectNativePivots(
  blob: Blob,
  data: IWorkbookData,
  pivots: PivotModel[] | undefined,
): Promise<Blob> {
  if (!nativePivotsEnabled()) return blob;
  const simple = (pivots ?? []).filter(isSimple);
  if (simple.length === 0) return blob;
  try {
    const zip = await JSZip.loadAsync(await blob.arrayBuffer());
    let n = maxPivotIndex(zip);
    for (const p of simple) {
      const src = readSource(data, p);
      if (!src || src.records.length === 0) continue;
      n += 1;
      const payload = generateNativePivot({
        sheetName: src.sheetName,
        sourceRef: `${cellRef(p.source.startRow, p.source.startColumn)}:${cellRef(p.source.endRow, p.source.endColumn)}`,
        headers: src.headers,
        records: src.records,
        rowField: p.rows[0].column,
        valueField: p.values[0].column,
        valueAgg: p.values[0].agg,
        targetRef: cellRef(p.target.row, p.target.column),
        index: n,
        cacheId: n,
      });
      await applyPivotsToZip(zip, payload);
    }
    return await zip.generateAsync({ type: 'blob' });
  } catch (err) {
    console.warn('[native-pivots] injection failed — exporting flat pivot cells', err);
    return blob;
  }
}
