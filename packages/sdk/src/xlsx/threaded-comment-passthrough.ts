import JSZip from 'jszip';

/**
 * Threaded-comment passthrough (authors + reply chains + timestamps).
 *
 * Modern Excel stores a comment thread in TWO layers:
 *   - the legacy `xl/comments{N}.xml` (+ VML) — a plain note, for back-compat;
 *   - `xl/threadedComments/threadedComment{N}.xml` — the real thread: each
 *     reply with its `personId`, ISO `dT`, and `parentId`, plus
 *     `xl/persons/person.xml` mapping `personId` → display name.
 *
 * ExcelJS only models the legacy note. Our bridge reads that note's text into a
 * Univer comment (author rewritten to "imported", replies + timestamps gone)
 * and writes it back as a legacy note on export — so the threaded layer is
 * silently dropped: every reply collapses to one note and every author is lost.
 *
 * Unlike images, threaded comments and persons aren't referenced by any element
 * in workbook/sheet XML — Excel discovers them purely by *relationship type*.
 * So this passthrough only has to restore the parts, declare their content
 * types, and re-create the relationships (workbook→persons, sheet→threadedComment);
 * no XML-element injection, no schema-order concerns. It rides on top of the
 * legacy note ExcelJS still writes (the required backing layer), so for an
 * open → save round-trip the full author/reply metadata survives in Excel.
 *
 * The parts are opaque bytes — we don't reconcile them with in-app comment
 * edits (those still flow through the legacy-note path).
 */

const REL_TYPE_PERSON = 'http://schemas.microsoft.com/office/2017/10/relationships/person';
const REL_TYPE_THREADED_COMMENT =
  'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment';
const CT_PERSON = 'application/vnd.ms-excel.person+xml';
const CT_THREADED_COMMENTS = 'application/vnd.ms-excel.threadedcomments+xml';

export interface ThreadedCommentPassthroughPayload {
  /** xl/threadedComments/** + xl/persons/** parts, keyed by zip path; base64. */
  parts: Record<string, string>;
  /** Per sheet (by decoded name): the threadedComment part it links to. */
  perSheet: Array<{ sheetName: string; target: string }>;
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

/** Resolve a rel Target relative to xl/worksheets/ (where sheet rels live). */
function resolveFromWorksheets(target: string): string {
  if (target.startsWith('/')) return target.replace(/^\//, '');
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
 * Capture every threaded-comment / persons part + the per-sheet threadedComment
 * linkage from a source xlsx. Returns undefined when there are no threaded
 * comments.
 */
export async function captureThreadedCommentsFromBuffer(
  buffer: ArrayBuffer,
): Promise<ThreadedCommentPassthroughPayload | undefined> {
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(buffer);
  } catch {
    return undefined;
  }

  const parts: Record<string, string> = {};
  for (const path of Object.keys(zip.files)) {
    if (zip.files[path].dir) continue;
    if (/^xl\/(threadedComments|persons)\//.test(path)) {
      parts[path] = await zip.files[path].async('base64');
    }
  }
  // Need at least one threadedComment part to be worth anything (a lone persons
  // part with no thread is meaningless).
  const hasThread = Object.keys(parts).some((p) => /^xl\/threadedComments\/[^/]+\.xml$/.test(p));
  if (!hasThread) return undefined;

  const perSheet: ThreadedCommentPassthroughPayload['perSheet'] = [];
  for (const sheet of await sheetEntries(zip)) {
    const relsXml = await zip.file(sheet.relsPath)?.async('string');
    if (!relsXml) continue;
    const rel = parseRels(relsXml).find((r) => r.type === REL_TYPE_THREADED_COMMENT);
    if (rel) {
      perSheet.push({ sheetName: sheet.name, target: resolveFromWorksheets(rel.target) });
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

const EMPTY_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';

/**
 * Re-inject captured threaded-comment + persons parts into the ExcelJS-written
 * zip and re-create their relationships. Mutates `zip` in place (composes with
 * the other passthrough steps). Sheets matched by decoded name.
 */
export async function applyThreadedCommentsToZip(
  zip: JSZip,
  payload: ThreadedCommentPassthroughPayload,
): Promise<void> {
  // 1. Restore every part verbatim.
  for (const [path, base64] of Object.entries(payload.parts)) {
    zip.file(path, base64, { base64: true });
  }

  // 2. [Content_Types].xml — Override per persons + threadedComments part.
  const ctEntry = zip.file('[Content_Types].xml');
  if (ctEntry) {
    let ct = await ctEntry.async('string');
    const inserts: string[] = [];
    for (const path of Object.keys(payload.parts)) {
      const isPerson = /^xl\/persons\/[^/]+\.xml$/.test(path);
      const isThread = /^xl\/threadedComments\/[^/]+\.xml$/.test(path);
      if (!isPerson && !isThread) continue;
      const partName = `/${path}`;
      if (ct.includes(`PartName="${partName}"`)) continue;
      const type = isPerson ? CT_PERSON : CT_THREADED_COMMENTS;
      inserts.push(`<Override PartName="${partName}" ContentType="${type}"/>`);
    }
    if (inserts.length > 0) {
      ct = ct.replace('</Types>', `${inserts.join('')}</Types>`);
      zip.file('[Content_Types].xml', ct);
    }
  }

  // 3. workbook.xml.rels — one person relationship per persons part (Excel
  //    discovers persons by relationship type, no workbook element needed).
  const personParts = Object.keys(payload.parts).filter((p) => /^xl\/persons\/[^/]+\.xml$/.test(p));
  if (personParts.length > 0) {
    const relsPath = 'xl/_rels/workbook.xml.rels';
    let relsXml = (await zip.file(relsPath)?.async('string')) ?? EMPTY_RELS;
    for (const part of personParts) {
      const target = part.replace(/^xl\//, ''); // relative to xl/
      if (relsXml.includes(`Target="${target}"`)) continue;
      const rId = nextRid(relsXml);
      const rel = `<Relationship Id="${rId}" Type="${REL_TYPE_PERSON}" Target="${target}"/>`;
      relsXml = relsXml.replace('</Relationships>', `${rel}</Relationships>`);
    }
    zip.file(relsPath, relsXml);
  }

  // 4. Per sheet: add the sheet→threadedComment relationship (fresh rId).
  const sheets = await sheetEntries(zip);
  const byName = new Map(sheets.map((s) => [s.name, s]));
  for (const { sheetName, target } of payload.perSheet) {
    const sheet = byName.get(sheetName);
    if (!sheet) continue; // sheet renamed/removed
    let relsXml = (await zip.file(sheet.relsPath)?.async('string')) ?? EMPTY_RELS;
    // worksheet-relative target (xl/threadedComments/x.xml → ../threadedComments/x.xml)
    const relTarget = `../${target.replace(/^xl\//, '')}`;
    if (relsXml.includes(`Target="${relTarget}"`)) continue; // already linked
    const rId = nextRid(relsXml);
    const rel = `<Relationship Id="${rId}" Type="${REL_TYPE_THREADED_COMMENT}" Target="${relTarget}"/>`;
    relsXml = relsXml.replace('</Relationships>', `${rel}</Relationships>`);
    zip.file(sheet.relsPath, relsXml);
  }
}
