import type ExcelJS from 'exceljs';
import type { IRange, IWorkbookData } from '@univerjs/core';

/**
 * xlsx-native `worksheet.dataValidations` ⇄ Univer
 * `SHEET_DATA_VALIDATION_PLUGIN` resource bridge.
 *
 * Data-validation rules live on the worksheet in xlsx (`xl/worksheets/
 * sheetN.xml` `<dataValidations>` element). Univer's data-validation
 * plugin keeps them in a resource keyed by sheet id under the name
 * below — same registration pattern as defined-names and
 * thread-comments. We mirror in both directions so:
 *
 *   - a file authored in real Excel keeps its list / whole / date /
 *     decimal constraints when opened here (parser fall-back)
 *   - our save round-trip preserves the rule even when the exporter
 *     is called without the live `extras` from the running app — eg.
 *     the audit harness drives `parse → snapshot → export` only.
 *
 * Resource name matches Univer's plugin (`vendor/univer/packages/
 * data-validation/src/controllers/dv-resource.controller.ts:23`) so
 * the rules live-load into the plugin on re-open instead of just
 * sitting in our sidecar — bonus on top of the round-trip.
 */

export const DATA_VALIDATION_RESOURCE = 'SHEET_DATA_VALIDATION_PLUGIN';

export type SynthDvRule = {
  uid: string;
  type: string; // 'list' | 'whole' | 'decimal' | 'date' | 'time' | 'textLength' | 'custom' | 'any'
  ranges: IRange[];
  formula1?: string;
  formula2?: string;
  operator?: string;
  allowBlank?: boolean;
  error?: string;
  errorTitle?: string;
  showErrorMessage?: boolean;
  errorStyle?: string;
  prompt?: string;
  promptTitle?: string;
  showInputMessage?: boolean;
};

function lettersToCol(letters: string): number {
  let col = 0;
  for (let i = 0; i < letters.length; i++) col = col * 26 + (letters.charCodeAt(i) - 64);
  return col - 1;
}

function colToLetters(n: number): string {
  let out = '';
  let v = n;
  while (v >= 0) {
    out = String.fromCharCode(65 + (v % 26)) + out;
    v = Math.floor(v / 26) - 1;
  }
  return out;
}

/**
 * Parse a single ExcelJS range token (`'D2:D4'` or `'$A$1'`) into an
 * IRange. Returns null on malformed input — callers should skip.
 */
function rangeStrToIRange(s: string): IRange | null {
  const part = s.includes('!') ? (s.split('!').pop() ?? s) : s;
  const m = /^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/.exec(part.trim());
  if (!m) return null;
  const c0 = lettersToCol(m[1]);
  const r0 = Number(m[2]) - 1;
  const c1 = m[3] ? lettersToCol(m[3]) : c0;
  const r1 = m[4] ? Number(m[4]) - 1 : r0;
  return { startRow: r0, endRow: r1, startColumn: c0, endColumn: c1 };
}

function iRangeToStr(r: IRange): string {
  const a = `${colToLetters(r.startColumn)}${r.startRow + 1}`;
  const b = `${colToLetters(r.endColumn)}${r.endRow + 1}`;
  return a === b ? a : `${a}:${b}`;
}

/**
 * Walk every worksheet and lift its `dataValidations.model` into the
 * synthesised plugin shape. ExcelJS stores rules keyed by a range
 * string that may be space-separated when a single rule covers
 * multiple ranges; we expand those into the `ranges` array.
 */
export function readDataValidationFromXlsx(
  wb: ExcelJS.Workbook,
  sheetIdForExcel: (excelId: number) => string,
): Record<string, SynthDvRule[]> {
  const out: Record<string, SynthDvRule[]> = {};
  let seq = 0;
  for (const ws of wb.worksheets) {
    const dv = (ws as unknown as { dataValidations?: { model?: Record<string, unknown> } })
      .dataValidations;
    const model = dv?.model;
    if (!model || typeof model !== 'object') continue;

    const rules: SynthDvRule[] = [];
    for (const [rangeStr, specUnknown] of Object.entries(model)) {
      const spec = specUnknown as Record<string, unknown>;
      if (!spec || typeof spec !== 'object') continue;
      const ranges: IRange[] = [];
      // ExcelJS keys can be `'D2:D4'`, `'D2:D4 F2:F4'`, or comma-joined.
      for (const piece of rangeStr.split(/[\s,]+/)) {
        const r = rangeStrToIRange(piece);
        if (r) ranges.push(r);
      }
      if (ranges.length === 0) continue;

      const formulae = Array.isArray(spec.formulae) ? (spec.formulae as unknown[]) : undefined;
      const rule: SynthDvRule = {
        uid: `dv-${seq++}`,
        type: typeof spec.type === 'string' ? spec.type : 'any',
        ranges,
        ...(typeof formulae?.[0] === 'string' ? { formula1: formulae[0] } : {}),
        ...(typeof formulae?.[1] === 'string' ? { formula2: formulae[1] } : {}),
        ...(typeof spec.operator === 'string' ? { operator: spec.operator } : {}),
        ...(typeof spec.allowBlank === 'boolean' ? { allowBlank: spec.allowBlank } : {}),
        ...(typeof spec.error === 'string' ? { error: spec.error } : {}),
        ...(typeof spec.errorTitle === 'string' ? { errorTitle: spec.errorTitle } : {}),
        ...(typeof spec.showErrorMessage === 'boolean'
          ? { showErrorMessage: spec.showErrorMessage }
          : {}),
        ...(typeof spec.errorStyle === 'string' ? { errorStyle: spec.errorStyle } : {}),
        ...(typeof spec.prompt === 'string' ? { prompt: spec.prompt } : {}),
        ...(typeof spec.promptTitle === 'string' ? { promptTitle: spec.promptTitle } : {}),
        ...(typeof spec.showInputMessage === 'boolean'
          ? { showInputMessage: spec.showInputMessage }
          : {}),
      };
      rules.push(rule);
    }
    if (rules.length > 0) out[sheetIdForExcel(ws.id)] = rules;
  }
  return out;
}

/**
 * Merge a synthesised DV map into the workbook's `resources` array.
 * Skipped when the workbook already carries the plugin resource (from
 * our hidden sidecar) — that one came through Univer's model and has
 * the full rule shape; xlsx-derived rules are a strict subset.
 */
export function mergeDataValidationIntoResources(
  resources: IWorkbookData['resources'],
  payload: Record<string, SynthDvRule[]>,
): IWorkbookData['resources'] {
  if (Object.keys(payload).length === 0) return resources;
  const existing = resources?.find((r) => r.name === DATA_VALIDATION_RESOURCE);
  if (existing) return resources;
  const next = [...(resources ?? [])];
  next.push({ name: DATA_VALIDATION_RESOURCE, data: JSON.stringify(payload) });
  return next;
}

/** Read the DV resource off a snapshot. Tolerant of older / missing /
 *  malformed payloads. */
export function readDataValidationFromSnapshot(
  data: IWorkbookData,
): Record<string, SynthDvRule[]> {
  const entry = data.resources?.find((r) => r.name === DATA_VALIDATION_RESOURCE);
  if (!entry?.data) return {};
  try {
    const parsed = JSON.parse(entry.data) as Record<string, unknown>;
    if (!parsed || typeof parsed !== 'object') return {};
    const out: Record<string, SynthDvRule[]> = {};
    for (const [sheetId, value] of Object.entries(parsed)) {
      if (!Array.isArray(value)) continue;
      const bucket: SynthDvRule[] = [];
      for (const r of value) {
        if (!r || typeof r !== 'object') continue;
        const obj = r as Record<string, unknown>;
        if (!Array.isArray(obj.ranges) || typeof obj.type !== 'string') continue;
        bucket.push(r as SynthDvRule);
      }
      if (bucket.length > 0) out[sheetId] = bucket;
    }
    return out;
  } catch {
    return {};
  }
}

/** Apply a per-sheet rule set onto an ExcelJS worksheet. */
export function applyDataValidationToXlsxWorksheet(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  ws: any,
  rules: SynthDvRule[],
): void {
  if (!rules?.length) return;
  for (const rule of rules) {
    const formulae: string[] = [];
    if (rule.formula1 !== undefined) formulae.push(rule.formula1);
    if (rule.formula2 !== undefined) formulae.push(rule.formula2);
    const spec: Record<string, unknown> = {
      type: rule.type,
      ...(formulae.length ? { formulae } : {}),
      ...(rule.operator ? { operator: rule.operator } : {}),
      ...(rule.allowBlank !== undefined ? { allowBlank: rule.allowBlank } : {}),
      ...(rule.error !== undefined ? { error: rule.error } : {}),
      ...(rule.errorTitle !== undefined ? { errorTitle: rule.errorTitle } : {}),
      ...(rule.showErrorMessage !== undefined ? { showErrorMessage: rule.showErrorMessage } : {}),
      ...(rule.errorStyle !== undefined ? { errorStyle: rule.errorStyle } : {}),
      ...(rule.prompt !== undefined ? { prompt: rule.prompt } : {}),
      ...(rule.promptTitle !== undefined ? { promptTitle: rule.promptTitle } : {}),
      ...(rule.showInputMessage !== undefined ? { showInputMessage: rule.showInputMessage } : {}),
    };
    for (const range of rule.ranges) {
      try {
        ws.dataValidations.add(iRangeToStr(range), spec);
      } catch {
        // Bad ranges (eg. references into a now-deleted sheet) shouldn't
        // kill the export. Drop the entry quietly; the rest still ship.
      }
    }
  }
}
