import JSZip from 'jszip';

/**
 * Raw-OOXML bridge for conditional-formatting rules that ExcelJS drops
 * entirely AND that carry a differential (`dxf`) style: today the
 * `duplicateValues` / `uniqueValues` highlight rules.
 *
 * ExcelJS never emits these `<cfRule>`s (no writer branch) and doesn't surface
 * them on read, so — like data bars — we own them end to end:
 *   - IMPORT: parse the `<cfRule type="duplicateValues|uniqueValues" dxfId=N>`
 *     out of the worksheet XML and resolve dxf N against `xl/styles.xml`'s
 *     `<dxfs>` into a fill/font style.
 *   - EXPORT: append a `<dxf>` to `<dxfs>` (creating the element if absent) and
 *     splice a `<cfRule>` referencing it into the worksheet XML, after ExcelJS
 *     has written the file.
 *
 * The differential-style subset matches the rest of the CF bridge: solid fill
 * background + font bold / italic / strike / colour.
 */

/** The CF fill/font style shared across the xlsx CF bridge (see
 *  conditional-formatting-resource.ts `CfStyle`). */
export interface CfStyle {
  bg?: { rgb: string };
  cl?: { rgb: string };
  bl?: number;
  it?: number;
  st?: { s: number };
}

export type DxfCfType = 'duplicateValues' | 'uniqueValues';

export interface DxfCfRule {
  type: DxfCfType;
  sqref: string;
  style: CfStyle;
}

const argbToHex = (argb: string | undefined): string | undefined => {
  if (typeof argb !== 'string') return undefined;
  const hex = /^([0-9A-Fa-f]{6,8})$/.exec(argb)?.[1];
  if (!hex) return undefined;
  return `#${(hex.length === 8 ? hex.slice(2) : hex).toLowerCase()}`;
};

const hexToArgb = (rgb: string | undefined): string | undefined => {
  const hex = typeof rgb === 'string' ? /^#?([0-9A-Fa-f]{6,8})$/.exec(rgb)?.[1] : undefined;
  if (!hex) return undefined;
  return (hex.length === 8 ? hex : `FF${hex}`).toUpperCase();
};

const xmlEscape = (s: string): string =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

const xmlUnescape = (s: string): string =>
  s
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(Number(d)))
    .replace(/&amp;/g, '&');

function parseRels(relsXml: string): Array<{ id: string; target: string }> {
  const out: Array<{ id: string; target: string }> = [];
  for (const m of relsXml.matchAll(/<Relationship\b[^>]*\/>/g)) {
    const id = m[0].match(/Id="([^"]+)"/)?.[1];
    const target = m[0].match(/Target="([^"]+)"/)?.[1];
    if (id && target) out.push({ id, target });
  }
  return out;
}

/** Map sheet name (decoded) → `xl/worksheets/sheetN.xml` path. */
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
    out.set(xmlUnescape(rawName), `xl/${target.replace(/^\/?xl\//, '')}`);
  }
  return out;
}

// --- dxf <-> CfStyle -------------------------------------------------------

/** Parse one `<dxf>...</dxf>` body into a CfStyle (fill bg + font traits). */
function parseDxf(dxfXml: string): CfStyle {
  const out: CfStyle = {};
  const fontBlock = dxfXml.match(/<font>([\s\S]*?)<\/font>/)?.[1] ?? '';
  if (/<b\/?>/.test(fontBlock)) out.bl = 1;
  if (/<i\/?>/.test(fontBlock)) out.it = 1;
  if (/<strike\/?>/.test(fontBlock)) out.st = { s: 1 };
  const fontColor = argbToHex(fontBlock.match(/<color\b[^>]*\brgb="([0-9A-Fa-f]{6,8})"/)?.[1]);
  if (fontColor) out.cl = { rgb: fontColor };
  // CF differential fills carry the highlight colour in bgColor.
  const fillBlock = dxfXml.match(/<fill>([\s\S]*?)<\/fill>/)?.[1] ?? '';
  const bg = argbToHex(
    fillBlock.match(/<bgColor\b[^>]*\brgb="([0-9A-Fa-f]{6,8})"/)?.[1] ??
      fillBlock.match(/<fgColor\b[^>]*\brgb="([0-9A-Fa-f]{6,8})"/)?.[1],
  );
  if (bg) out.bg = { rgb: bg };
  return out;
}

/** Serialise a CfStyle to a `<dxf>` element (mirrors ExcelJS's dxf shape). */
function dxfXmlForStyle(style: CfStyle): string {
  let font = '';
  if (style.bl === 1) font += '<b/>';
  if (style.it === 1) font += '<i/>';
  if (style.st?.s === 1) font += '<strike/>';
  const cl = hexToArgb(style.cl?.rgb);
  if (cl) font += `<color rgb="${cl}"/>`;
  const fontXml = font ? `<font>${font}</font>` : '';
  const bg = hexToArgb(style.bg?.rgb);
  const fillXml = bg
    ? `<fill><patternFill patternType="solid"><bgColor rgb="${bg}"/></patternFill></fill>`
    : '';
  return `<dxf>${fontXml}${fillXml}</dxf>`;
}

// --- IMPORT ----------------------------------------------------------------

/** Parse the `<dxfs>` element of styles.xml into an ordered CfStyle array. */
function parseDxfs(stylesXml: string): CfStyle[] {
  const block = stylesXml.match(/<dxfs\b[^>]*>([\s\S]*?)<\/dxfs>/)?.[1];
  if (!block) return [];
  return [...block.matchAll(/<dxf>[\s\S]*?<\/dxf>|<dxf\/>/g)].map((m) => parseDxf(m[0]));
}

/** Pull duplicate/unique cfRules + their resolved dxf style from one sheet XML. */
function readDxfCfRules(sheetXml: string, dxfs: CfStyle[]): DxfCfRule[] {
  const out: DxfCfRule[] = [];
  for (const block of sheetXml.matchAll(
    /<conditionalFormatting\b[^>]*sqref="([^"]+)"[^>]*>([\s\S]*?)<\/conditionalFormatting>/g,
  )) {
    const sqref = xmlUnescape(block[1]);
    for (const rule of block[2].matchAll(/<cfRule\b[^>]*\/>|<cfRule\b[^>]*>[\s\S]*?<\/cfRule>/g)) {
      const tag = rule[0];
      const type = tag.match(/\btype="(duplicateValues|uniqueValues)"/)?.[1] as
        | DxfCfType
        | undefined;
      if (!type) continue;
      const dxfId = Number(tag.match(/\bdxfId="(\d+)"/)?.[1]);
      const style = Number.isNaN(dxfId) ? {} : (dxfs[dxfId] ?? {});
      out.push({ type, sqref, style });
    }
  }
  return out;
}

/**
 * Recover duplicate/unique CF rules from a source xlsx buffer, keyed by sheet
 * name. Returns undefined when there are none (so callers can skip).
 */
export async function captureDxfCfRulesFromBuffer(
  buffer: ArrayBuffer,
): Promise<Record<string, DxfCfRule[]> | undefined> {
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(buffer);
  } catch {
    return undefined;
  }
  const stylesXml = await zip.file('xl/styles.xml')?.async('string');
  const dxfs = stylesXml ? parseDxfs(stylesXml) : [];
  const paths = await sheetNameToPath(zip);
  const out: Record<string, DxfCfRule[]> = {};
  for (const [name, path] of paths) {
    const xml = await zip.file(path)?.async('string');
    if (!xml || !/type="(duplicateValues|uniqueValues)"/.test(xml)) continue;
    const rules = readDxfCfRules(xml, dxfs);
    if (rules.length > 0) out[name] = rules;
  }
  return Object.keys(out).length > 0 ? out : undefined;
}

// --- EXPORT ----------------------------------------------------------------

// Elements that follow <conditionalFormatting> in CT_Worksheet schema order.
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

// Elements that follow <dxfs> in CT_Stylesheet schema order.
const AFTER_DXFS = ['<tableStyles', '<colors', '<extLst', '</styleSheet>'];

/** Append dxfs to styles.xml (creating <dxfs> if absent); returns the base
 *  index the new dxfs start at, and the patched XML. */
function appendDxfs(stylesXml: string, styles: CfStyle[]): { xml: string; baseIndex: number } {
  const newDxfs = styles.map(dxfXmlForStyle).join('');
  const existing = stylesXml.match(/<dxfs\b[^>]*>([\s\S]*?)<\/dxfs>/);
  if (existing) {
    // dxfId is a positional index into the <dxf> children, so base our indices
    // on the actual child count (not the `count` attr, which a foreign file
    // could disagree with).
    const baseIndex = (existing[1].match(/<dxf[\s>/]/g) ?? []).length;
    const merged = `<dxfs count="${baseIndex + styles.length}">${existing[1]}${newDxfs}</dxfs>`;
    return { xml: stylesXml.replace(existing[0], merged), baseIndex };
  }
  // Self-closing or absent <dxfs> — insert a fresh element in schema order.
  const selfClosed = stylesXml.match(/<dxfs\b[^>]*\/>/);
  const block = `<dxfs count="${styles.length}">${newDxfs}</dxfs>`;
  if (selfClosed) return { xml: stylesXml.replace(selfClosed[0], block), baseIndex: 0 };
  for (const tag of AFTER_DXFS) {
    const idx = stylesXml.indexOf(tag);
    if (idx !== -1) {
      return { xml: stylesXml.slice(0, idx) + block + stylesXml.slice(idx), baseIndex: 0 };
    }
  }
  return { xml: stylesXml, baseIndex: 0 };
}

function spliceCfRules(
  sheetXml: string,
  rules: Array<{ type: DxfCfType; sqref: string; dxfId: number }>,
): string {
  if (rules.length === 0) return sheetXml;
  let maxPriority = 0;
  for (const m of sheetXml.matchAll(/priority="(\d+)"/g))
    maxPriority = Math.max(maxPriority, +m[1]);
  const blocks = rules
    .map(
      (r, i) =>
        `<conditionalFormatting sqref="${xmlEscape(r.sqref)}">` +
        `<cfRule type="${r.type}" dxfId="${r.dxfId}" priority="${maxPriority + 1 + i}"/>` +
        `</conditionalFormatting>`,
    )
    .join('');
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
 * Inject every duplicate/unique CF rule into the ExcelJS-written zip: one shared
 * `<dxf>` per rule in styles.xml, and the `<cfRule>` spliced into its sheet.
 * Mutates `zip` in place; sheets matched by decoded name.
 */
export async function applyDxfCfRulesToZip(
  zip: JSZip,
  rulesBySheet: Record<string, DxfCfRule[]>,
): Promise<void> {
  const paths = await sheetNameToPath(zip);

  // 1. Collect rules for sheets that still exist, in a stable order, and assign
  //    each a global dxf index appended to styles.xml.
  const planned: Array<{ path: string; type: DxfCfType; sqref: string; style: CfStyle }> = [];
  for (const [name, rules] of Object.entries(rulesBySheet)) {
    const path = paths.get(name);
    if (!path) continue; // sheet renamed/removed
    for (const r of rules) planned.push({ path, type: r.type, sqref: r.sqref, style: r.style });
  }
  if (planned.length === 0) return;

  const stylesEntry = zip.file('xl/styles.xml');
  if (!stylesEntry) return;
  const { xml: stylesXml, baseIndex } = appendDxfs(
    await stylesEntry.async('string'),
    planned.map((p) => p.style),
  );
  zip.file('xl/styles.xml', stylesXml);

  // 2. Splice cfRules per sheet, referencing the dxf indices assigned above.
  const byPath = new Map<string, Array<{ type: DxfCfType; sqref: string; dxfId: number }>>();
  planned.forEach((p, i) => {
    const list = byPath.get(p.path) ?? [];
    list.push({ type: p.type, sqref: p.sqref, dxfId: baseIndex + i });
    byPath.set(p.path, list);
  });
  for (const [path, rules] of byPath) {
    const xml = await zip.file(path)?.async('string');
    if (!xml) continue;
    zip.file(path, spliceCfRules(xml, rules));
  }
}
