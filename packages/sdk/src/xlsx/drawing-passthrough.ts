import JSZip from 'jszip';

/**
 * Drawing (image / shape) passthrough.
 *
 * ExcelJS rebuilds the exported workbook from the Univer snapshot, which has no
 * drawing model — so every embedded image and shape is **silently dropped** on
 * save (the parser never reads `xl/media` / `xl/drawings` either). For anyone
 * trading real `.xlsx`, that's data loss with no warning.
 *
 * This captures the raw drawing OOXML at parse time and re-injects it at export
 * (mirrors `pivot-passthrough.ts`). An embedded image lives across several parts:
 *   - `xl/media/image{N}.{ext}`            — the binary
 *   - `xl/drawings/drawing{N}.xml`         — anchors (which cells the image spans)
 *   - `xl/drawings/_rels/drawing{N}.xml.rels` — drawing → media link
 *   - `xl/worksheets/_rels/sheet{N}.xml.rels` — sheet → drawing link
 *   - `<drawing r:id="..."/>` in `sheet{N}.xml`
 *
 * We ship `xl/media/**` + `xl/drawings/**` verbatim (their internal rels point at
 * media by relative path, which never changes), then on export re-create the
 * sheet→drawing link with a fresh rId on the ExcelJS-regenerated sheet. Sheets
 * are matched by NAME (ExcelJS may reorder files). Images aren't rendered in the
 * editor yet — this round-trips them so Excel keeps them. Shapes/SmartArt ride
 * along for free (same parts).
 */

const REL_TYPE_DRAWING =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing';
const CT_DRAWING = 'application/vnd.openxmlformats-officedocument.drawing+xml';

// Image extension → content-type for the [Content_Types].xml Default entries.
const IMAGE_CONTENT_TYPES: Record<string, string> = {
  png: 'image/png',
  jpeg: 'image/jpeg',
  jpg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  tif: 'image/tiff',
  tiff: 'image/tiff',
  emf: 'image/x-emf',
  wmf: 'image/x-wmf',
  svg: 'image/svg+xml',
  webp: 'image/webp',
};

export interface DrawingPassthroughPayload {
  /** xl/media/** + xl/drawings/** parts, keyed by zip path; base64 contents. */
  parts: Record<string, string>;
  /** Per sheet (by name): the drawing part it links to (e.g. xl/drawings/drawing1.xml). */
  perSheet: Array<{ sheetName: string; drawingTarget: string }>;
}

function parseRels(xml: string): Array<{ id: string; type: string; target: string }> {
  const out: Array<{ id: string; type: string; target: string }> = [];
  for (const m of xml.matchAll(/<Relationship\b[^>]*\/>/g)) {
    const tag = m[0];
    const id = tag.match(/Id="([^"]+)"/)?.[1];
    const type = tag.match(/Type="([^"]+)"/)?.[1];
    const target = tag.match(/Target="([^"]+)"/)?.[1];
    if (id && type && target) out.push({ id, type, target });
  }
  return out;
}

const xmlUnescape = (s: string): string =>
  s
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(Number(d)))
    .replace(/&amp;/g, '&');

/** Resolve a relationship Target against `xl/worksheets/` (where sheet rels
 *  live) to an absolute zip path. Handles `../drawings/x.xml` and absolute. */
function resolveFromWorksheets(target: string): string {
  if (target.startsWith('/')) return target.replace(/^\//, '');
  // sheet rels are relative to xl/worksheets/
  const stack = 'xl/worksheets'.split('/');
  for (const seg of target.split('/')) {
    if (seg === '..') stack.pop();
    else if (seg !== '.') stack.push(seg);
  }
  return stack.join('/');
}

/** Map decoded sheet name → { sheetPath, relsPath } from a zip. */
async function sheetEntries(
  zip: JSZip,
): Promise<Array<{ name: string; sheetPath: string; relsPath: string }>> {
  const out: Array<{ name: string; sheetPath: string; relsPath: string }> = [];
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
    const sheetPath = `xl/${target.replace(/^\/?xl\//, '')}`;
    const file = sheetPath.split('/').pop()!;
    out.push({
      name: xmlUnescape(rawName),
      sheetPath,
      relsPath: `xl/worksheets/_rels/${file}.rels`,
    });
  }
  return out;
}

/**
 * Capture every drawing/image part from a source xlsx buffer, plus the
 * sheet→drawing linkage. Returns undefined when the file has no drawings.
 */
export async function captureDrawingsFromBuffer(
  buffer: ArrayBuffer,
): Promise<DrawingPassthroughPayload | undefined> {
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(buffer);
  } catch {
    return undefined;
  }
  const parts: Record<string, string> = {};
  for (const path of Object.keys(zip.files)) {
    if (zip.files[path].dir) continue;
    if (/^xl\/(media|drawings)\//.test(path)) {
      parts[path] = await zip.files[path].async('base64');
    }
  }
  if (Object.keys(parts).length === 0) return undefined;

  const perSheet: DrawingPassthroughPayload['perSheet'] = [];
  for (const sheet of await sheetEntries(zip)) {
    const relsXml = await zip.file(sheet.relsPath)?.async('string');
    if (!relsXml) continue;
    const drawingRel = parseRels(relsXml).find((r) => r.type === REL_TYPE_DRAWING);
    if (drawingRel) {
      perSheet.push({
        sheetName: sheet.name,
        drawingTarget: resolveFromWorksheets(drawingRel.target),
      });
    }
  }
  return { parts, perSheet };
}

function nextRid(relsXml: string): string {
  const used = new Set<number>();
  for (const m of relsXml.matchAll(/Id="rId(\d+)"/g)) used.add(Number(m[1]));
  let n = 1;
  while (used.has(n)) n++;
  return `rId${n}`;
}

// Worksheet elements that must follow <drawing> in CT_Worksheet schema order;
// we insert the <drawing> just before the first one present.
const AFTER_DRAWING = [
  '<legacyDrawing',
  '<legacyDrawingHF',
  '<drawingHF',
  '<picture',
  '<oleObjects',
  '<controls',
  '<webPublishItems',
  '<tableParts',
  '<extLst',
  '</worksheet>',
];

/**
 * Re-inject captured drawing parts into the ExcelJS-written zip and re-link each
 * sheet's `<drawing>`. Mutates `zip` in place (composes with VBA/pivot/CF
 * passthrough). Sheets matched by decoded name.
 */
export async function applyDrawingsToZip(
  zip: JSZip,
  payload: DrawingPassthroughPayload,
): Promise<void> {
  // 1. Restore every media + drawing part verbatim.
  for (const [path, base64] of Object.entries(payload.parts)) {
    zip.file(path, base64, { base64: true });
  }

  // 2. [Content_Types].xml — image Defaults (per extension present) + a drawing
  //    Override per drawing part. ExcelJS omits these (no images in the model).
  const ctEntry = zip.file('[Content_Types].xml');
  if (ctEntry) {
    let ct = await ctEntry.async('string');
    const inserts: string[] = [];
    const exts = new Set<string>();
    for (const path of Object.keys(payload.parts)) {
      const m = /^xl\/media\/[^.]+\.([A-Za-z0-9]+)$/.exec(path);
      if (m) exts.add(m[1].toLowerCase());
    }
    for (const ext of exts) {
      const type = IMAGE_CONTENT_TYPES[ext] ?? `image/${ext}`;
      if (!new RegExp(`<Default Extension="${ext}"`, 'i').test(ct)) {
        inserts.push(`<Default Extension="${ext}" ContentType="${type}"/>`);
      }
    }
    for (const path of Object.keys(payload.parts)) {
      if (!/^xl\/drawings\/drawing[^/]+\.xml$/.test(path)) continue;
      const partName = `/${path}`;
      if (ct.includes(`PartName="${partName}"`)) continue;
      inserts.push(`<Override PartName="${partName}" ContentType="${CT_DRAWING}"/>`);
    }
    if (inserts.length > 0) {
      ct = ct.replace('</Types>', `${inserts.join('')}</Types>`);
      zip.file('[Content_Types].xml', ct);
    }
  }

  // 3. Per sheet: add the sheet→drawing relationship (fresh rId) + inject the
  //    <drawing r:id> element into the regenerated sheet XML.
  const sheets = await sheetEntries(zip);
  const byName = new Map(sheets.map((s) => [s.name, s]));
  for (const { sheetName, drawingTarget } of payload.perSheet) {
    const sheet = byName.get(sheetName);
    if (!sheet) continue; // sheet renamed/removed
    const sheetXml = await zip.file(sheet.sheetPath)?.async('string');
    if (!sheetXml || sheetXml.includes('<drawing ')) continue; // already linked

    // drawing target as a worksheet-relative path (xl/drawings/x.xml → ../drawings/x.xml)
    const relTarget = `../${drawingTarget.replace(/^xl\//, '')}`;
    let relsXml =
      (await zip.file(sheet.relsPath)?.async('string')) ??
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
    const rId = nextRid(relsXml);
    const rel = `<Relationship Id="${rId}" Type="${REL_TYPE_DRAWING}" Target="${relTarget}"/>`;
    relsXml = relsXml.includes('</Relationships>')
      ? relsXml.replace('</Relationships>', `${rel}</Relationships>`)
      : `${relsXml}\n${rel}`;
    zip.file(sheet.relsPath, relsXml);

    const drawingEl = `<drawing r:id="${rId}"/>`;
    let patched = sheetXml;
    const idx = AFTER_DRAWING.map((t) => patched.indexOf(t)).filter((i) => i !== -1);
    const at = idx.length ? Math.min(...idx) : -1;
    patched = at !== -1 ? patched.slice(0, at) + drawingEl + patched.slice(at) : patched;
    zip.file(sheet.sheetPath, patched);
  }
}
