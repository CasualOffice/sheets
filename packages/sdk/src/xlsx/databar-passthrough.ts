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

import JSZip from 'jszip';

/**
 * Data-bar conditional-formatting raw-OOXML bridge.
 *
 * ExcelJS surfaces almost every data-bar field on read (cfvo, gradient, lengths,
 * axis/direction/border flags) but drops the one that matters most — the
 * positive bar **fill colour** — and on write it emits a broken `<color
 * auto="1"/>` with no fill. So the normal ExcelJS path can't carry a data bar
 * faithfully. We therefore:
 *   - IMPORT: read the positive `<color rgb>` straight out of the worksheet XML
 *     here (everything else comes from ExcelJS), and build a Univer IDataBar.
 *   - EXPORT: skip ExcelJS for data bars entirely and splice a correct legacy
 *     `<conditionalFormatting><cfRule type="dataBar">` block into the worksheet
 *     XML after ExcelJS has written the file (mirrors `pivot-passthrough.ts`).
 *
 * Scope: the legacy block (positive colour + min/max + showValue). The x14
 * extension (explicit gradient flag, negative/axis colours) is out — Excel
 * renders a sensible gradient bar from the legacy block, and axis/border/
 * direction have no slot in Univer's IDataBar anyway.
 */

/** A threshold stop — Univer IValueConfig / OOXML cfvo (shared shape). */
interface ValueConfig {
  type: string;
  value?: number | string;
}

/** Everything needed to re-emit one data bar on export. */
export interface DataBarEntry {
  sqref: string;
  positiveColor: string; // '#rrggbb'
  isShowValue: boolean;
  min: ValueConfig;
  max: ValueConfig;
}

/** Positive fill colour recovered from the worksheet XML, keyed by range. */
export interface DataBarColor {
  sqref: string;
  positiveColor: string; // '#rrggbb'
}

const argbToHex = (argb: string | undefined): string | undefined => {
  if (typeof argb !== 'string') return undefined;
  const hex = /^([0-9A-Fa-f]{6,8})$/.exec(argb)?.[1];
  if (!hex) return undefined;
  return `#${(hex.length === 8 ? hex.slice(2) : hex).toLowerCase()}`;
};

const hexToArgb = (rgb: string | undefined): string => {
  const hex = typeof rgb === 'string' ? /^#?([0-9A-Fa-f]{6,8})$/.exec(rgb)?.[1] : undefined;
  if (!hex) return 'FF638EC6'; // Excel's default data-bar blue
  return (hex.length === 8 ? hex : `FF${hex}`).toUpperCase();
};

const xmlEscape = (s: string): string =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/** Decode XML entities in an attribute value (workbook.xml stores sheet names
 *  escaped, but ExcelJS/Univer report them decoded — keys must match those). */
const xmlUnescape = (s: string): string =>
  s
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(Number(d)))
    .replace(/&amp;/g, '&'); // last, so "&amp;lt;" → "&lt;" not "<"

/** Parse `<Relationship Id Type Target/>` entries (paths are workbook-relative). */
function parseRels(relsXml: string): Array<{ id: string; target: string }> {
  const out: Array<{ id: string; target: string }> = [];
  for (const m of relsXml.matchAll(/<Relationship\b[^>]*\/>/g)) {
    const id = m[0].match(/Id="([^"]+)"/)?.[1];
    const target = m[0].match(/Target="([^"]+)"/)?.[1];
    if (id && target) out.push({ id, target });
  }
  return out;
}

/** Build a map of sheet name → `xl/worksheets/sheetN.xml` path from a zip. */
async function sheetNameToPath(zip: JSZip): Promise<Map<string, string>> {
  const out = new Map<string, string>();
  const workbookXml = await zip.file('xl/workbook.xml')?.async('string');
  const relsXml = await zip.file('xl/_rels/workbook.xml.rels')?.async('string');
  if (!workbookXml || !relsXml) return out;
  const rels = parseRels(relsXml);
  for (const m of workbookXml.matchAll(/<sheet\b[^>]*\/>/g)) {
    const rawName = m[0].match(/name="([^"]+)"/)?.[1];
    const rId = m[0].match(/r:id="([^"]+)"/i)?.[1];
    if (!rawName || !rId) continue;
    const target = rels.find((r) => r.id === rId)?.target;
    if (!target) continue;
    // Key by the DECODED name to match ExcelJS/Univer (which report sheet names
    // decoded); workbook.xml stores them XML-escaped.
    // Targets are workbook-relative (e.g. "worksheets/sheet1.xml").
    out.set(xmlUnescape(rawName), `xl/${target.replace(/^\/?xl\//, '')}`);
  }
  return out;
}

/** Pull each data bar's positive `<color rgb>` out of one worksheet's XML. */
export function readDataBarColors(sheetXml: string): DataBarColor[] {
  const out: DataBarColor[] = [];
  for (const block of sheetXml.matchAll(
    /<conditionalFormatting\b[^>]*sqref="([^"]+)"[^>]*>([\s\S]*?)<\/conditionalFormatting>/g,
  )) {
    const sqref = block[1];
    // Within the block, each dataBar's positive colour is the <color> child of
    // <dataBar> (after the two <cfvo>). Take the first rgb colour we find.
    for (const db of block[2].matchAll(/<dataBar\b[^>]*>([\s\S]*?)<\/dataBar>/g)) {
      const rgb = db[1].match(/<color\b[^>]*\brgb="([0-9A-Fa-f]{6,8})"/)?.[1];
      const positiveColor = argbToHex(rgb);
      if (positiveColor) out.push({ sqref, positiveColor });
    }
  }
  return out;
}

/**
 * Recover data-bar positive colours from a source xlsx buffer, keyed by sheet
 * name. Returns undefined when the file has no data bars (so callers can skip).
 */
export async function captureDataBarColorsFromBuffer(
  buffer: ArrayBuffer,
): Promise<Record<string, DataBarColor[]> | undefined> {
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(buffer);
  } catch {
    return undefined;
  }
  const paths = await sheetNameToPath(zip);
  const out: Record<string, DataBarColor[]> = {};
  for (const [name, path] of paths) {
    const xml = await zip.file(path)?.async('string');
    if (!xml || !xml.includes('type="dataBar"')) continue;
    const colors = readDataBarColors(xml);
    if (colors.length > 0) out[name] = colors;
  }
  return Object.keys(out).length > 0 ? out : undefined;
}

function cfvoXml(vc: ValueConfig): string {
  if (vc.type === 'min' || vc.type === 'max') return `<cfvo type="${vc.type}"/>`;
  const val = Number(vc.value) || 0;
  return `<cfvo type="${vc.type}" val="${val}"/>`;
}

function dataBarBlockXml(entry: DataBarEntry, priority: number): string {
  return (
    `<conditionalFormatting sqref="${xmlEscape(entry.sqref)}">` +
    `<cfRule type="dataBar" priority="${priority}">` +
    `<dataBar showValue="${entry.isShowValue ? 1 : 0}">` +
    cfvoXml(entry.min) +
    cfvoXml(entry.max) +
    `<color rgb="${hexToArgb(entry.positiveColor)}"/>` +
    `</dataBar></cfRule></conditionalFormatting>`
  );
}

// Worksheet child elements that must follow <conditionalFormatting> per the
// CT_Worksheet schema; we insert our block just before the first one present.
const AFTER_CF = [
  '<dataValidations',
  '<hyperlinks',
  '<printOptions',
  '<pageMargins',
  '<pageSetup',
  '<headerFooter',
  '<rowBreaks',
  '<colBreaks',
  '<drawing',
  '<legacyDrawing',
  '<tableParts',
  '<extLst',
  '</worksheet>',
];

/** Splice generated data-bar blocks into the ExcelJS-written worksheet XML. */
function spliceDataBars(sheetXml: string, entries: DataBarEntry[]): string {
  if (entries.length === 0) return sheetXml;
  // Next free priority = max existing + 1 (data bars and other CF coexist).
  let maxPriority = 0;
  for (const m of sheetXml.matchAll(/priority="(\d+)"/g)) {
    maxPriority = Math.max(maxPriority, Number(m[1]));
  }
  const blocks = entries.map((e, i) => dataBarBlockXml(e, maxPriority + 1 + i)).join('');

  // Prefer to place after the last existing <conditionalFormatting> (keeps all
  // CF together, schema-valid); else before the first trailing element.
  const lastCfClose = sheetXml.lastIndexOf('</conditionalFormatting>');
  if (lastCfClose !== -1) {
    const at = lastCfClose + '</conditionalFormatting>'.length;
    return sheetXml.slice(0, at) + blocks + sheetXml.slice(at);
  }
  for (const tag of AFTER_CF) {
    const idx = sheetXml.indexOf(tag);
    if (idx !== -1) return sheetXml.slice(0, idx) + blocks + sheetXml.slice(idx);
  }
  return sheetXml;
}

/**
 * Inject every captured data bar into the ExcelJS-written zip. Mutates `zip`
 * in place; the caller finalises it (so this composes with VBA/pivot
 * passthrough). Sheets are matched by NAME against the regenerated workbook.
 */
export async function applyDataBarsToZip(
  zip: JSZip,
  dataBars: Record<string, DataBarEntry[]>,
): Promise<void> {
  const paths = await sheetNameToPath(zip);
  for (const [name, entries] of Object.entries(dataBars)) {
    const path = paths.get(name);
    if (!path) continue; // sheet renamed or removed — skip
    const xml = await zip.file(path)?.async('string');
    if (!xml) continue;
    zip.file(path, spliceDataBars(xml, entries));
  }
}
