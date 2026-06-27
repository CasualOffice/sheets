import type ExcelJS from 'exceljs';
import type { IRange, IWorkbookData } from '@univerjs/core';
import type { DataBarEntry } from './databar-passthrough';
import type { DxfCfRule } from './cf-dxf-passthrough';

/**
 * xlsx-native `worksheet.conditionalFormattings` ⇄ Univer
 * `SHEET_CONDITIONAL_FORMATTING_PLUGIN` resource bridge.
 *
 * Conditional-formatting rules live inline in `xl/worksheets/sheetN.xml`
 * (`<conditionalFormatting>`); ExcelJS surfaces them as
 * `worksheet.conditionalFormattings` (`[{ ref, rules }]`). Univer's CF plugin
 * keeps them in a resource keyed by sheet id under the name below (serialized as
 * `{ [sheetId]: IConditionFormattingRule[] }`, see the fork's
 * conditional-formatting.service.ts `toJson`). We mirror in both directions so a
 * file authored in real Excel keeps its highlight rules when opened here, and our
 * save round-trip preserves them.
 *
 * Scope: **highlight-cell** rules — mapped to Univer's IHighlightCell subtypes:
 *   - `number`   ← ExcelJS `cellIs` (numeric comparisons)
 *   - `formula`  ← ExcelJS `expression`
 *   - `rank`     ← ExcelJS `top10` (top/bottom N, optionally percent)
 *   - `average`  ← ExcelJS `aboveAverage` (above / below the range mean)
 *   - `timePeriod` ← ExcelJS `timePeriod` (today / last7Days / thisMonth / …)
 *   - `text`     ← ExcelJS `containsText` (operators: containsText, plus the
 *                  no-value blanks/errors operators)
 * each with the rule's fill / font style; the `duplicateValues` /
 * `uniqueValues` highlight rules (style only); plus the visual rule types
 *   - `colorScale` ← ExcelJS `colorScale` (value-mapped gradient stops)
 *   - `iconSet`    ← ExcelJS `iconSet` (named icon group + threshold bands)
 *   - `dataBar`    ← ExcelJS `dataBar` (value-proportional bar)
 * which have no fill/font style. (A colorScale/iconSet using a `formula`
 * threshold is dropped — ExcelJS floatifies a cfvo value on read, destroying
 * the formula text.) ExcelJS can't carry `dataBar` or `duplicateValues` /
 * `uniqueValues` itself (it drops the rule and/or its colour), so those are
 * bridged via raw OOXML — see databar-passthrough.ts and cf-dxf-passthrough.ts.
 *
 * Deliberately NOT mapped (ExcelJS itself can't round-trip them, so they'd be
 * silently lost rather than preserved — see the bridge tests): the text
 * `beginsWith` / `endsWith` / `notContainsText` operators (ExcelJS drops their
 * search text). Unmapped rules are skipped, never corrupted.
 */

export const CONDITIONAL_FORMATTING_RESOURCE = 'SHEET_CONDITIONAL_FORMATTING_PLUGIN';

// Numeric comparison operators shared verbatim between ExcelJS `cellIs` and
// Univer's CFNumberOperator (same string values).
const NUMBER_OPERATORS = new Set([
  'greaterThan',
  'greaterThanOrEqual',
  'lessThan',
  'lessThanOrEqual',
  'between',
  'notBetween',
  'equal',
  'notEqual',
]);

// ExcelJS `containsText` operators we preserve. `containsText` carries a search
// string (recovered from its formula); the rest are value-less predicates.
// `beginsWith` / `endsWith` / `notContainsText` are intentionally absent —
// ExcelJS drops their search text on round-trip, leaving a meaningless rule.
const TEXT_VALUE_OPERATOR = 'containsText';
const TEXT_VALUELESS_OPERATORS = new Set([
  'containsBlanks',
  'notContainsBlanks',
  'containsErrors',
  'notContainsErrors',
]);

// ExcelJS `timePeriod` values map verbatim to Univer's CFTimePeriodOperator.
const TIME_PERIODS = new Set([
  'today',
  'yesterday',
  'tomorrow',
  'last7Days',
  'thisMonth',
  'lastMonth',
  'nextMonth',
  'thisWeek',
  'lastWeek',
  'nextWeek',
]);

// ExcelJS color-scale / data-bar / icon-set thresholds (`cfvo`) carry a type +
// optional value. Univer's CFValueType is the same set minus ExcelJS's
// auto-min/max (which collapse to plain min/max). `formula` is intentionally
// absent: ExcelJS parses a cfvo's `val` with parseFloat on read, so a formula
// threshold's text is destroyed (becomes NaN) — it can't round-trip, so a rule
// using one is dropped rather than emitted with a corrupt "NaN" stop.
const CFVO_TYPE_TO_UNIVER: Record<string, string> = {
  num: 'num',
  percent: 'percent',
  percentile: 'percentile',
  min: 'min',
  max: 'max',
  autoMin: 'min',
  autoMax: 'max',
};
const UNIVER_VALUE_TYPES = new Set(['num', 'percent', 'percentile', 'min', 'max']);

/** A Univer IValueConfig — a conditional-formatting threshold stop. */
interface CfValueConfig {
  type: string;
  value?: number | string;
}

/** ExcelJS `cfvo` entry → Univer IValueConfig, or null if untranslatable. */
function cfvoToValueConfig(cfvo: unknown): CfValueConfig | null {
  const c = cfvo as { type?: string; value?: unknown };
  const type = c?.type ? CFVO_TYPE_TO_UNIVER[c.type] : undefined;
  if (!type) return null;
  if (type === 'min' || type === 'max') return { type };
  const n = Number(c.value);
  if (Number.isNaN(n)) return null;
  return { type, value: n };
}

/** Univer IValueConfig → ExcelJS `cfvo` entry. */
function valueConfigToCfvo(vc: CfValueConfig): Record<string, unknown> {
  if (vc.type === 'min' || vc.type === 'max') return { type: vc.type };
  return { type: vc.type, value: Number(vc.value) || 0 };
}

function isValueConfig(v: unknown): v is CfValueConfig {
  const c = v as CfValueConfig;
  return !!c && typeof c.type === 'string' && UNIVER_VALUE_TYPES.has(c.type);
}

// Icon-set names we round-trip: the standard ECMA-376 `ST_IconSetType` values,
// which are exactly Univer's IIconSetType minus its non-OOXML extras. The three
// Excel-2010 x14 sets (`3Triangles` / `3Stars` / `5Boxes`) are excluded — ExcelJS
// treats them as non-primitive and its base `<iconSet>` writer drops them (they
// belong in an x14 extension it doesn't emit), so claiming them would lose the
// rule on save. Such a rule is skipped instead.
const ICON_SET_TYPES = new Set([
  '3Arrows',
  '3ArrowsGray',
  '4Arrows',
  '4ArrowsGray',
  '5Arrows',
  '5ArrowsGray',
  '3TrafficLights1',
  '3TrafficLights2',
  '3Signs',
  '4RedToBlack',
  '4TrafficLights',
  '3Symbols',
  '3Symbols2',
  '3Flags',
  '4Rating',
  '5Rating',
  '5Quarters',
]);

interface CfStyle {
  bg?: { rgb: string };
  cl?: { rgb: string };
  bl?: number;
  it?: number;
  st?: { s: number };
}

// `rule` mirrors Univer's IHighlightCell subtypes. A discriminated union keeps
// each subtype's required fields honest both when synthesising from xlsx and
// when reading back off a snapshot for export.
type SynthRule =
  | {
      type: 'highlightCell';
      subType: 'number';
      operator: string;
      value: number | [number, number];
      style: CfStyle;
    }
  | { type: 'highlightCell'; subType: 'formula'; value: string; style: CfStyle }
  | { type: 'highlightCell'; subType: 'text'; operator: string; value?: string; style: CfStyle }
  | {
      type: 'highlightCell';
      subType: 'rank';
      isBottom: boolean;
      isPercent: boolean;
      value: number;
      style: CfStyle;
    }
  | { type: 'highlightCell'; subType: 'average'; operator: string; style: CfStyle }
  | { type: 'highlightCell'; subType: 'timePeriod'; operator: string; style: CfStyle }
  // duplicate/unique highlight rules carry only a style (no operator/value).
  // ExcelJS drops them entirely, so they're bridged via raw XML — see
  // cf-dxf-passthrough.ts.
  | { type: 'highlightCell'; subType: 'duplicateValues' | 'uniqueValues'; style: CfStyle }
  // Visual rule. colorScale has no fill/font style — it paints a value-mapped
  // gradient; `config` is the ordered list of gradient stops.
  | { type: 'colorScale'; config: Array<{ index: number; color: string; value: CfValueConfig }> }
  // Visual rule. iconSet paints a per-cell icon. `config` is Univer-ordered
  // (descending: config[0] = highest band, iconId '0'); no fill/font style.
  | {
      type: 'iconSet';
      isShowValue: boolean;
      config: Array<{ operator: string; value: CfValueConfig; iconType: string; iconId: string }>;
    }
  // Visual rule. dataBar paints a value-proportional bar; no fill/font style.
  // The positive bar colour is recovered from raw XML on import and re-emitted
  // via raw XML on export (ExcelJS botches it) — see databar-passthrough.ts.
  | {
      type: 'dataBar';
      isShowValue: boolean;
      config: {
        min: CfValueConfig;
        max: CfValueConfig;
        isGradient: boolean;
        positiveColor: string;
        nativeColor: string;
      };
    };

interface SynthCfRule {
  cfId: string;
  ranges: IRange[];
  stopIfTrue: boolean;
  rule: SynthRule;
}

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

function normalizeArgb(argb: string | undefined): string | undefined {
  if (typeof argb !== 'string') return undefined;
  const hex = /^#?([0-9A-Fa-f]{6,8})$/.exec(argb)?.[1];
  if (!hex) return undefined;
  return `#${(hex.length === 8 ? hex.slice(2) : hex).toLowerCase()}`;
}

function toArgb(rgb: string | undefined): string | undefined {
  if (typeof rgb !== 'string') return undefined;
  const hex = /^#?([0-9A-Fa-f]{6,8})$/.exec(rgb)?.[1];
  if (!hex) return undefined;
  return (hex.length === 8 ? hex : `FF${hex}`).toUpperCase();
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function dxfToCfStyle(style: any): CfStyle {
  const out: CfStyle = {};
  // CF differential fills carry the highlight colour in bgColor (the OOXML
  // dxf patternFill convention); fall back to fgColor just in case.
  const bg = normalizeArgb(style?.fill?.bgColor?.argb) ?? normalizeArgb(style?.fill?.fgColor?.argb);
  if (bg) out.bg = { rgb: bg };
  const cl = normalizeArgb(style?.font?.color?.argb);
  if (cl) out.cl = { rgb: cl };
  if (style?.font?.bold) out.bl = 1;
  if (style?.font?.italic) out.it = 1;
  if (style?.font?.strike) out.st = { s: 1 };
  return out;
}

function cfStyleToDxf(rawStyle: CfStyle | undefined | null): Record<string, unknown> {
  // A foreign / partially-formed resource payload may carry a rule with no
  // style; tolerate it rather than throwing and aborting the whole export.
  const style = rawStyle ?? {};
  const out: Record<string, unknown> = {};
  const font: Record<string, unknown> = {};
  if (style.bl === 1) font.bold = true;
  if (style.it === 1) font.italic = true;
  if (style.st?.s === 1) font.strike = true;
  const cl = toArgb(style.cl?.rgb);
  if (cl) font.color = { argb: cl };
  if (Object.keys(font).length) out.font = font;
  const bg = toArgb(style.bg?.rgb);
  if (bg)
    out.fill = { type: 'pattern', pattern: 'solid', bgColor: { argb: bg }, fgColor: { argb: bg } };
  return out;
}

/** Recover a `containsText` rule's search string from its formula. ExcelJS drops
 *  the `text` attribute on read, leaving only the formula it (and Excel) emit:
 *  `NOT(ISERROR(SEARCH("term",A1)))`. Pull the first SEARCH("…") literal,
 *  un-escaping Excel's doubled quotes. Returns undefined when no match. */
function containsTextValue(formulae: unknown[]): string | undefined {
  const f = typeof formulae[0] === 'string' ? formulae[0] : '';
  const m = /SEARCH\("((?:[^"]|"")*)"/.exec(f);
  return m ? m[1].replace(/""/g, '"') : undefined;
}

/** Map one ExcelJS conditional-formatting rule to a Univer highlight-cell rule,
 *  or null if its type/operator isn't one we preserve (see the module header). */
function excelRuleToSynthRule(raw: unknown, injectedColor?: string): SynthRule | null {
  const r = raw as {
    type?: string;
    operator?: string;
    timePeriod?: string;
    rank?: number;
    percent?: boolean;
    bottom?: boolean;
    aboveAverage?: boolean;
    text?: string;
    formulae?: unknown[];
    cfvo?: unknown[];
    color?: unknown[];
    iconSet?: string;
    reverse?: boolean;
    showValue?: boolean;
    gradient?: boolean;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    style?: any;
  };
  const style = dxfToCfStyle(r.style);
  const formulae = Array.isArray(r.formulae) ? r.formulae : [];

  if (r.type === 'cellIs' && typeof r.operator === 'string' && NUMBER_OPERATORS.has(r.operator)) {
    const n0 = Number(formulae[0]);
    if (Number.isNaN(n0)) return null; // non-numeric cellIs — not a number rule
    const between = r.operator === 'between' || r.operator === 'notBetween';
    const n1 = Number(formulae[1]);
    if (between && Number.isNaN(n1)) return null;
    return {
      type: 'highlightCell',
      subType: 'number',
      operator: r.operator,
      value: between ? [n0, n1] : n0,
      style,
    };
  }

  if (r.type === 'expression' && typeof formulae[0] === 'string') {
    return { type: 'highlightCell', subType: 'formula', value: formulae[0], style };
  }

  if (r.type === 'top10') {
    const value = Number(r.rank);
    if (Number.isNaN(value)) return null;
    return {
      type: 'highlightCell',
      subType: 'rank',
      isBottom: r.bottom === true,
      isPercent: r.percent === true,
      value,
      style,
    };
  }

  if (r.type === 'aboveAverage') {
    // ExcelJS omits the attribute when above-average (its default), so only an
    // explicit `false` means below-average.
    return {
      type: 'highlightCell',
      subType: 'average',
      operator: r.aboveAverage === false ? 'lessThan' : 'greaterThan',
      style,
    };
  }

  if (
    r.type === 'timePeriod' &&
    typeof r.timePeriod === 'string' &&
    TIME_PERIODS.has(r.timePeriod)
  ) {
    return { type: 'highlightCell', subType: 'timePeriod', operator: r.timePeriod, style };
  }

  // ExcelJS folds containsText / blanks / errors under type `containsText`,
  // discriminated by `operator`. (beginsWith / endsWith / notContainsText keep
  // their own type and lose their text — we don't map those.)
  if (r.type === 'containsText' && typeof r.operator === 'string') {
    if (r.operator === TEXT_VALUE_OPERATOR) {
      // A loaded xlsx drops the `text` attribute (recover from the formula); an
      // in-memory ExcelJS rule has `text` but no formula yet. Prefer whichever
      // is present.
      const value = typeof r.text === 'string' ? r.text : containsTextValue(formulae);
      if (value === undefined) return null; // couldn't recover the search text
      return { type: 'highlightCell', subType: 'text', operator: r.operator, value, style };
    }
    if (TEXT_VALUELESS_OPERATORS.has(r.operator)) {
      return { type: 'highlightCell', subType: 'text', operator: r.operator, style };
    }
  }

  // Color scale: parallel `cfvo` (threshold stops) + `color` arrays. Univer
  // keeps an ordered `config` of { index, color, value } gradient stops.
  if (r.type === 'colorScale' && Array.isArray(r.cfvo) && Array.isArray(r.color)) {
    const n = Math.min(r.cfvo.length, r.color.length);
    if (n < 2) return null; // a gradient needs at least two stops
    const config: Array<{ index: number; color: string; value: CfValueConfig }> = [];
    for (let i = 0; i < n; i++) {
      const value = cfvoToValueConfig(r.cfvo[i]);
      const color = normalizeArgb((r.color[i] as { argb?: string })?.argb);
      if (!value || !color) return null; // bail rather than emit a broken gradient
      config.push({ index: i, color, value });
    }
    return { type: 'colorScale', config };
  }

  // Icon set: a named icon group + `cfvo` threshold stops (ascending in OOXML).
  // Univer keeps the bands DESCENDING (config[0] = highest), iconId = position
  // in the icon group. `reverse` flips which icon maps to which band.
  if (
    r.type === 'iconSet' &&
    typeof r.iconSet === 'string' &&
    ICON_SET_TYPES.has(r.iconSet) &&
    Array.isArray(r.cfvo)
  ) {
    const n = r.cfvo.length;
    if (n < 2) return null;
    const stops = r.cfvo.map(cfvoToValueConfig);
    if (stops.some((s) => !s)) return null; // a formula/garbled threshold — skip
    const reverse = r.reverse === true;
    const config = stops.map((_stop, i) => {
      const last = i === n - 1;
      return {
        // Non-last bands use >= their stop; the last band is the unconditional
        // fallback (Univer renders it regardless), mirroring the canonical shape.
        operator: last ? 'lessThanOrEqual' : 'greaterThanOrEqual',
        value: stops[n - 1 - i]!, // invert ascending cfvo → descending bands
        iconType: r.iconSet!,
        iconId: String(reverse ? n - 1 - i : i),
      };
    });
    return { type: 'iconSet', isShowValue: r.showValue !== false, config };
  }

  // Data bar: ExcelJS gives the shape (cfvo / gradient) but never the colour —
  // the positive fill is recovered from raw XML by the caller and threaded in
  // as `injectedColor` (see databar-passthrough.ts). Negative colour has no
  // ExcelJS surface either, so default it.
  if (r.type === 'dataBar' && Array.isArray(r.cfvo) && r.cfvo.length >= 2) {
    const min = cfvoToValueConfig(r.cfvo[0]);
    const max = cfvoToValueConfig(r.cfvo[1]);
    if (!min || !max) return null; // formula/garbled anchor — skip
    return {
      type: 'dataBar',
      isShowValue: r.showValue !== false,
      config: {
        min,
        max,
        isGradient: r.gradient !== false,
        positiveColor: injectedColor ?? '#638ec6', // Excel's default data-bar blue
        nativeColor: '#ff0000',
      },
    };
  }

  return null;
}

const normalizeSqref = (s: string): string => s.replace(/\s+/g, ' ').trim().toUpperCase();

/** Lift every worksheet's conditional formatting into the synthesised plugin
 *  shape. `dataBarColors` carries the positive bar colours recovered from raw
 *  XML (keyed by sheet name), since ExcelJS drops them — see
 *  databar-passthrough.ts. */
export function readConditionalFormattingFromXlsx(
  wb: ExcelJS.Workbook,
  sheetIdForExcel: (excelId: number) => string,
  dataBarColors?: Record<string, Array<{ sqref: string; positiveColor: string }>>,
): Record<string, SynthCfRule[]> {
  const out: Record<string, SynthCfRule[]> = {};
  let seq = 0;
  for (const ws of wb.worksheets) {
    const cfs = (ws as unknown as { conditionalFormattings?: unknown }).conditionalFormattings;
    if (!Array.isArray(cfs)) continue;

    const sheetColors = dataBarColors?.[ws.name] ?? [];

    const rules: SynthCfRule[] = [];
    for (const cf of cfs) {
      const ref = (cf as { ref?: string }).ref;
      const cfRules = (cf as { rules?: unknown[] }).rules;
      if (typeof ref !== 'string' || !Array.isArray(cfRules)) continue;
      const ranges: IRange[] = [];
      for (const piece of ref.split(/[\s,]+/)) {
        const r = rangeStrToIRange(piece);
        if (r) ranges.push(r);
      }
      if (ranges.length === 0) continue;

      for (const raw of cfRules) {
        const isDataBar = (raw as { type?: string }).type === 'dataBar';
        const injectedColor = isDataBar
          ? sheetColors.find((c) => normalizeSqref(c.sqref) === normalizeSqref(ref))?.positiveColor
          : undefined;
        const rule = excelRuleToSynthRule(raw, injectedColor);
        if (rule) rules.push({ cfId: `cf-${seq++}`, ranges, stopIfTrue: false, rule });
        // Unmapped rule types (see the module header) are skipped, not corrupted.
      }
    }
    if (rules.length > 0) out[sheetIdForExcel(ws.id)] = rules;
  }
  return out;
}

/** Merge a synthesised CF map into the workbook's `resources` array. */
export function mergeConditionalFormattingIntoResources(
  resources: IWorkbookData['resources'],
  payload: Record<string, SynthCfRule[]>,
): IWorkbookData['resources'] {
  if (Object.keys(payload).length === 0) return resources;
  // Prefer an existing plugin resource (came through Univer's live model with
  // the full rule shape) over our xlsx-derived subset.
  const existing = resources?.find((r) => r.name === CONDITIONAL_FORMATTING_RESOURCE);
  if (existing) return resources;
  const next = [...(resources ?? [])];
  next.push({ name: CONDITIONAL_FORMATTING_RESOURCE, data: JSON.stringify(payload) });
  return next;
}

/** Read the CF resource off a snapshot, keeping only number/formula highlight
 *  rules (the shapes we can export). Tolerant of missing/foreign payloads. */
export function readConditionalFormattingFromSnapshot(
  data: IWorkbookData,
): Record<string, SynthCfRule[]> {
  const entry = data.resources?.find((r) => r.name === CONDITIONAL_FORMATTING_RESOURCE);
  if (!entry?.data) return {};
  try {
    const parsed = JSON.parse(entry.data) as Record<string, unknown>;
    if (!parsed || typeof parsed !== 'object') return {};
    const out: Record<string, SynthCfRule[]> = {};
    for (const [sheetId, value] of Object.entries(parsed)) {
      if (!Array.isArray(value)) continue;
      const bucket: SynthCfRule[] = [];
      for (const item of value) {
        const obj = item as SynthCfRule;
        if (obj && Array.isArray(obj.ranges) && isExportableSynthRule(obj.rule)) {
          bucket.push(obj);
        }
      }
      if (bucket.length > 0) out[sheetId] = bucket;
    }
    return out;
  } catch {
    return {};
  }
}

/** Read data-bar rules off a snapshot's CF resource as raw-XML export entries,
 *  keyed by sheetId. ExcelJS can't write a data bar, so the export pipeline
 *  emits these via databar-passthrough.ts instead. */
export function readDataBarsFromSnapshot(data: IWorkbookData): Record<string, DataBarEntry[]> {
  const entry = data.resources?.find((r) => r.name === CONDITIONAL_FORMATTING_RESOURCE);
  if (!entry?.data) return {};
  let parsed: Record<string, unknown>;
  try {
    parsed = JSON.parse(entry.data) as Record<string, unknown>;
  } catch {
    return {};
  }
  if (!parsed || typeof parsed !== 'object') return {};
  const out: Record<string, DataBarEntry[]> = {};
  for (const [sheetId, value] of Object.entries(parsed)) {
    if (!Array.isArray(value)) continue;
    const bars: DataBarEntry[] = [];
    for (const item of value) {
      const obj = item as SynthCfRule;
      if (!obj || !Array.isArray(obj.ranges) || obj.rule?.type !== 'dataBar') continue;
      if (!isExportableSynthRule(obj.rule)) continue;
      const sqref = obj.ranges.map(iRangeToStr).join(' ');
      if (!sqref) continue;
      bars.push({
        sqref,
        positiveColor: obj.rule.config.positiveColor,
        isShowValue: obj.rule.isShowValue !== false,
        min: obj.rule.config.min,
        max: obj.rule.config.max,
      });
    }
    if (bars.length > 0) out[sheetId] = bars;
  }
  return out;
}

/** Read duplicate/unique rules off a snapshot's CF resource as raw-XML export
 *  entries, keyed by sheetId. ExcelJS can't write them, so the export pipeline
 *  emits these via cf-dxf-passthrough.ts. */
export function readDxfCfRulesFromSnapshot(data: IWorkbookData): Record<string, DxfCfRule[]> {
  const entry = data.resources?.find((r) => r.name === CONDITIONAL_FORMATTING_RESOURCE);
  if (!entry?.data) return {};
  let parsed: Record<string, unknown>;
  try {
    parsed = JSON.parse(entry.data) as Record<string, unknown>;
  } catch {
    return {};
  }
  if (!parsed || typeof parsed !== 'object') return {};
  const out: Record<string, DxfCfRule[]> = {};
  for (const [sheetId, value] of Object.entries(parsed)) {
    if (!Array.isArray(value)) continue;
    const rules: DxfCfRule[] = [];
    for (const item of value) {
      const obj = item as SynthCfRule;
      if (!obj || !Array.isArray(obj.ranges) || obj.rule?.type !== 'highlightCell') continue;
      if (obj.rule.subType !== 'duplicateValues' && obj.rule.subType !== 'uniqueValues') continue;
      const sqref = obj.ranges.map(iRangeToStr).join(' ');
      if (!sqref) continue;
      rules.push({ type: obj.rule.subType, sqref, style: obj.rule.style ?? {} });
    }
    if (rules.length > 0) out[sheetId] = rules;
  }
  return out;
}

/** Convert duplicate/unique rules captured from raw XML (keyed by sheetId) into
 *  synthesised CF rules to merge into the resource map. */
export function dxfCfRulesToSynthCf(
  rulesBySheetId: Record<string, DxfCfRule[]>,
): Record<string, SynthCfRule[]> {
  const out: Record<string, SynthCfRule[]> = {};
  let seq = 0;
  for (const [sheetId, rules] of Object.entries(rulesBySheetId)) {
    const bucket: SynthCfRule[] = [];
    for (const r of rules) {
      const ranges: IRange[] = [];
      for (const piece of r.sqref.split(/[\s,]+/)) {
        const range = rangeStrToIRange(piece);
        if (range) ranges.push(range);
      }
      if (ranges.length === 0) continue;
      bucket.push({
        cfId: `cf-dxf-${seq++}`,
        ranges,
        stopIfTrue: false,
        rule: { type: 'highlightCell', subType: r.type, style: r.style },
      });
    }
    if (bucket.length > 0) out[sheetId] = bucket;
  }
  return out;
}

/** Validate a synthesised rule we can faithfully export (guards foreign /
 *  partially-mapped payloads read off a snapshot). */
function isExportableSynthRule(rule: SynthCfRule['rule'] | undefined): rule is SynthRule {
  if (!rule) return false;
  if (rule.type === 'dataBar') {
    // Kept so the rule survives the snapshot read (renders on re-open) and is
    // available to the raw-XML export path; synthRuleToExcel deliberately
    // refuses it so ExcelJS never writes a (broken) data bar.
    const c = rule.config;
    return (
      !!c && isValueConfig(c.min) && isValueConfig(c.max) && typeof c.positiveColor === 'string'
    );
  }
  if (rule.type === 'colorScale') {
    return (
      Array.isArray(rule.config) &&
      rule.config.length >= 2 &&
      rule.config.every((c) => typeof c.color === 'string' && isValueConfig(c.value))
    );
  }
  if (rule.type === 'iconSet') {
    return (
      Array.isArray(rule.config) &&
      rule.config.length >= 2 &&
      rule.config.every(
        (c) =>
          typeof c.iconType === 'string' &&
          ICON_SET_TYPES.has(c.iconType) &&
          typeof c.iconId === 'string' &&
          typeof c.operator === 'string' &&
          isValueConfig(c.value),
      )
    );
  }
  if (rule.type !== 'highlightCell') return false;
  switch (rule.subType) {
    case 'number':
      return typeof rule.operator === 'string' && rule.value !== undefined;
    case 'formula':
      return typeof rule.value === 'string';
    case 'rank':
      return typeof rule.value === 'number' && !Number.isNaN(rule.value);
    case 'average':
      return typeof rule.operator === 'string';
    case 'timePeriod':
      return typeof rule.operator === 'string' && TIME_PERIODS.has(rule.operator);
    case 'text':
      return rule.operator === TEXT_VALUE_OPERATOR
        ? typeof rule.value === 'string'
        : TEXT_VALUELESS_OPERATORS.has(rule.operator);
    case 'duplicateValues':
    case 'uniqueValues':
      // Kept so it survives the snapshot read (renders + feeds the raw-XML
      // export); synthRuleToExcel refuses it so ExcelJS never writes one.
      return true;
    default:
      return false;
  }
}

/** Turn one synthesised rule into the ExcelJS shape that round-trips back to the
 *  same rule. Returns null for shapes ExcelJS can't faithfully write. */
function synthRuleToExcel(rule: SynthRule, priority: number): Record<string, unknown> | null {
  // Data bars are written via raw XML (databar-passthrough.ts), never ExcelJS —
  // ExcelJS emits a broken `<color auto="1"/>` with no fill.
  if (rule.type === 'dataBar') return null;
  // duplicate/unique are written via raw XML (cf-dxf-passthrough.ts) — ExcelJS
  // has no writer branch for them.
  if (
    rule.type === 'highlightCell' &&
    (rule.subType === 'duplicateValues' || rule.subType === 'uniqueValues')
  ) {
    return null;
  }
  if (rule.type === 'colorScale') {
    const ordered = [...rule.config].sort((a, b) => a.index - b.index);
    return {
      type: 'colorScale',
      priority,
      cfvo: ordered.map((c) => valueConfigToCfvo(c.value)),
      color: ordered.map((c) => ({ argb: toArgb(c.color) })),
    };
  }
  if (rule.type === 'iconSet') {
    const n = rule.config.length;
    // OOXML cfvo is ascending; Univer config is descending → reverse back.
    const cfvo = rule.config.map((_c, j) => valueConfigToCfvo(rule.config[n - 1 - j].value));
    // A non-'0' first iconId means the import inverted the icon order (reverse).
    const reverse = rule.config[0]?.iconId !== '0';
    return {
      type: 'iconSet',
      iconSet: rule.config[0].iconType,
      cfvo,
      showValue: rule.isShowValue,
      reverse,
      priority,
    };
  }
  const style = cfStyleToDxf(rule.style);
  switch (rule.subType) {
    case 'number': {
      const formulae = Array.isArray(rule.value) ? rule.value.map(String) : [String(rule.value)];
      return { type: 'cellIs', operator: rule.operator, formulae, priority, style };
    }
    case 'formula':
      return { type: 'expression', formulae: [String(rule.value)], priority, style };
    case 'rank':
      return {
        type: 'top10',
        rank: rule.value,
        percent: rule.isPercent,
        bottom: rule.isBottom,
        priority,
        style,
      };
    case 'average':
      // Univer's average operator is greater/less-than the mean; ExcelJS models
      // it as the boolean `aboveAverage`.
      return {
        type: 'aboveAverage',
        aboveAverage: rule.operator === 'greaterThan' || rule.operator === 'greaterThanOrEqual',
        priority,
        style,
      };
    case 'timePeriod':
      return { type: 'timePeriod', timePeriod: rule.operator, priority, style };
    case 'text':
      // ExcelJS folds these under `containsText`, building the formula from
      // `text` for the value operator; the rest are value-less.
      return {
        type: 'containsText',
        operator: rule.operator,
        ...(rule.operator === TEXT_VALUE_OPERATOR ? { text: rule.value ?? '' } : {}),
        priority,
        style,
      };
    default:
      return null;
  }
}

/** Apply a per-sheet CF rule set onto an ExcelJS worksheet. */
export function applyConditionalFormattingToXlsxWorksheet(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  ws: any,
  rules: SynthCfRule[],
): void {
  if (!rules?.length) return;
  rules.forEach((entry, i) => {
    const ref = entry.ranges.map(iRangeToStr).join(' ');
    if (!ref) return;
    if (!isExportableSynthRule(entry.rule)) return;
    const exceljsRule = synthRuleToExcel(entry.rule, i + 1);
    if (!exceljsRule) return;
    try {
      ws.addConditionalFormatting({ ref, rules: [exceljsRule] });
    } catch {
      // A malformed rule shouldn't kill the export; drop it quietly.
    }
  });
}
