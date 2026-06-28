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
 * Complex pivot cache + table passthrough.
 *
 * ExcelJS reads xlsx at the worksheet API level and never touches the
 * pivot OOXML parts. On a naive open→save loop, every pivot
 * definition gets silently dropped — the materialised cells survive,
 * but Excel sees the file as flat data on re-open.
 *
 * This module captures the raw OOXML parts at parse time and re-
 * injects them at export. Compared to VBA passthrough (single binary
 * with one rel + one Content_Type entry), pivots require:
 *
 *  1. Multiple parts: `xl/pivotCaches/pivotCacheDefinition{N}.xml`,
 *     `xl/pivotCaches/pivotCacheRecords{N}.xml`,
 *     `xl/pivotTables/pivotTable{N}.xml`, plus each part's own
 *     `.rels` (which we ship verbatim — those rId spaces are local).
 *  2. Rel renumbering across `xl/_rels/workbook.xml.rels` AND every
 *     sheet's `xl/worksheets/_rels/sheet{N}.xml.rels`. Original rIds
 *     would collide with ExcelJS-assigned ones, so we map old → new
 *     at inject time.
 *  3. Surgery on the ExcelJS-regenerated `xl/workbook.xml` to inject
 *     a `<pivotCaches>` element with the REMAPPED rIds, since
 *     ExcelJS regenerates that file from scratch and drops the
 *     element.
 *  4. Multiple `Override` entries in `[Content_Types].xml` for the
 *     three pivot ContentTypes.
 *
 * Sheet matching at inject time: the original capture records pivots
 * by **sheet name** (not by sheet rId), because ExcelJS may reorder
 * sheet files. At inject we look up the new sheet path via the
 * regenerated workbook.xml `<sheets>` block, then write into that
 * sheet's `.rels`.
 *
 * Constraints/assumptions:
 *  - Univer renders the materialised cells from sheet.xml; the
 *    pivot table itself is dropped from Univer's runtime model
 *    today. Edits to the underlying data don't propagate back into
 *    the pivot cache. This passthrough only preserves the OOXML so
 *    Excel re-recognises the file as having pivots — not full
 *    interactive pivot editing in our app.
 *  - We never modify the pivot parts' own contents.
 *  - We don't try to reconcile the cached pivot data with the live
 *    cells if the user edits the source range. Excel's "Refresh"
 *    handles that on its end.
 */

/** Three OOXML namespaces that identify pivot parts in `*.rels`. */
const REL_TYPE_PIVOT_CACHE_DEF =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition';
const REL_TYPE_PIVOT_TABLE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable';

const CT_PIVOT_CACHE_DEF =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml';
const CT_PIVOT_CACHE_REC =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml';
const CT_PIVOT_TABLE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml';

export type PivotPassthroughPayload = {
  /** Every captured part — pivotCaches/**, pivotTables/**, and any
   *  per-part `.rels` files — keyed by zip path; base64 contents. */
  parts: Record<string, string>;
  /** The raw `<pivotCaches>…</pivotCaches>` block extracted from
   *  `xl/workbook.xml` at parse time. The `r:id` refs inside get
   *  remapped at inject time. */
  workbookPivotCachesXml: string;
  /** Workbook-level pivotCacheDefinition relationships. `origId` is
   *  the captured rId — used to remap the `r:id` references inside
   *  `workbookPivotCachesXml`. */
  workbookCacheRels: Array<{ origId: string; target: string }>;
  /** Pivot tables, grouped by the sheet they belong to. Sheets are
   *  identified by NAME (not rId) because ExcelJS may reorder files. */
  perSheet: Array<{
    sheetName: string;
    /** sheet's pivotTable rels. `origId` is the captured rId in the
     *  ORIGINAL sheet rels file; we assign fresh rIds at inject time. */
    pivotTableRels: Array<{ origId: string; target: string }>;
  }>;
};

/** Best-effort namespace-tolerant `r:id="..."` extractor. */
function extractRid(xml: string): string | null {
  const m = xml.match(/r:id="([^"]+)"/i) ?? xml.match(/r:Id="([^"]+)"/i);
  return m ? m[1] : null;
}

/** Strip the wrapping <Relationships> shell — used so we can append
 *  captured rels without doubling up. */
function parseRelationships(xml: string): Array<{ id: string; type: string; target: string }> {
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

/** Closure that returns the next free `rIdN` string each call, given
 *  a Set tracking which numbers are already used. The set is mutated
 *  in place so successive calls don't collide. */
function nextRidIssuer(used: Set<number>): () => string {
  return () => {
    let n = 1;
    while (used.has(n)) n++;
    used.add(n);
    return `rId${n}`;
  };
}

/**
 * Walk the input buffer (the raw uploaded xlsx) and pull out every
 * piece of pivot machinery. Returns `undefined` when no pivot parts
 * are present — callers should skip the rest of the pipeline in that
 * case to avoid bloating snapshots with empty payloads.
 */
export async function capturePivotsFromBuffer(
  buffer: ArrayBuffer,
): Promise<PivotPassthroughPayload | undefined> {
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(buffer);
  } catch {
    return undefined;
  }

  // Quick presence check — is there anything pivot-shaped at all?
  const hasPivots = Object.keys(zip.files).some(
    (p) => p.startsWith('xl/pivotCaches/') || p.startsWith('xl/pivotTables/'),
  );
  if (!hasPivots) return undefined;

  // 1. Capture every pivot-cache and pivot-table part + their own .rels.
  const parts: Record<string, string> = {};
  for (const [path, file] of Object.entries(zip.files)) {
    if (file.dir) continue;
    if (
      path.startsWith('xl/pivotCaches/') ||
      path.startsWith('xl/pivotTables/')
    ) {
      parts[path] = await file.async('base64');
    }
  }

  // 2. Extract <pivotCaches>...</pivotCaches> from xl/workbook.xml.
  const workbookXmlEntry = zip.file('xl/workbook.xml');
  if (!workbookXmlEntry) return undefined;
  const workbookXml = await workbookXmlEntry.async('string');
  const pivotCachesMatch = workbookXml.match(/<pivotCaches\b[^>]*>[\s\S]*?<\/pivotCaches>/);
  if (!pivotCachesMatch) return undefined;
  const workbookPivotCachesXml = pivotCachesMatch[0];

  // 3. Workbook-level pivotCacheDefinition rels.
  const workbookRelsEntry = zip.file('xl/_rels/workbook.xml.rels');
  if (!workbookRelsEntry) return undefined;
  const workbookRelsXml = await workbookRelsEntry.async('string');
  const workbookCacheRels = parseRelationships(workbookRelsXml)
    .filter((r) => r.type === REL_TYPE_PIVOT_CACHE_DEF)
    .map((r) => ({ origId: r.id, target: r.target }));

  // 4. Per-sheet pivotTable rels. Need to map sheet rels file → sheet
  //    name via workbook.xml + workbook.xml.rels.
  // workbook.xml: <sheet name="..." r:id="rIdN"/>
  // workbook.xml.rels: rIdN → Target="worksheets/sheetN.xml"
  // sheetN.xml.rels: pivot table rels live here.
  const sheetEntries: Array<{ name: string; rId: string; targetPath: string }> = [];
  for (const m of workbookXml.matchAll(/<sheet\b[^>]*\/>/g)) {
    const tag = m[0];
    const name = tag.match(/name="([^"]+)"/)?.[1];
    const rId = tag.match(/r:id="([^"]+)"/i)?.[1] ?? tag.match(/r:Id="([^"]+)"/)?.[1];
    if (!name || !rId) continue;
    const rel = parseRelationships(workbookRelsXml).find((r) => r.id === rId);
    if (!rel) continue;
    sheetEntries.push({ name, rId, targetPath: rel.target });
  }

  const perSheet: PivotPassthroughPayload['perSheet'] = [];
  for (const sheet of sheetEntries) {
    // target like "worksheets/sheet1.xml" — rels live at
    // "xl/worksheets/_rels/sheet1.xml.rels".
    const fileName = sheet.targetPath.split('/').pop()!;
    const sheetRelsPath = `xl/worksheets/_rels/${fileName}.rels`;
    const sheetRelsEntry = zip.file(sheetRelsPath);
    if (!sheetRelsEntry) continue;
    const sheetRelsXml = await sheetRelsEntry.async('string');
    const pivotTableRels = parseRelationships(sheetRelsXml)
      .filter((r) => r.type === REL_TYPE_PIVOT_TABLE)
      .map((r) => ({ origId: r.id, target: r.target }));
    if (pivotTableRels.length > 0) {
      perSheet.push({ sheetName: sheet.name, pivotTableRels });
    }
  }

  // Only emit a payload when we actually have something to round-trip.
  if (
    Object.keys(parts).length === 0 ||
    workbookCacheRels.length === 0
  ) {
    return undefined;
  }

  return {
    parts,
    workbookPivotCachesXml,
    workbookCacheRels,
    perSheet,
  };
}

/**
 * Re-inject every captured pivot piece into the ExcelJS-written xlsx
 * buffer. The input buffer is read but not mutated; the same JSZip
 * instance is finalised by the caller (so caller can compose multiple
 * passthrough steps).
 *
 * Mutates the passed zip in-place. Caller does the final
 * generateAsync. This shape lets `passthrough-resource.ts` compose
 * VBA and pivot passthrough into one final zip write.
 */
export async function applyPivotsToZip(
  zip: JSZip,
  payload: PivotPassthroughPayload,
): Promise<void> {
  // 1. Drop every captured part back at its original path.
  for (const [path, base64] of Object.entries(payload.parts)) {
    zip.file(path, base64, { base64: true });
  }

  // 2. Patch [Content_Types].xml — add Overrides for any pivot part
  //    that doesn't already have one. We pick the ContentType from
  //    the path: pivotCacheDefinition / pivotCacheRecords / pivotTable.
  const ctEntry = zip.file('[Content_Types].xml');
  if (ctEntry) {
    let ct = await ctEntry.async('string');
    const inserts: string[] = [];
    for (const path of Object.keys(payload.parts)) {
      // Skip .rels files — those don't need ContentType Overrides
      // (the default rels ContentType in Defaults covers them).
      if (path.endsWith('.rels')) continue;
      const partName = `/${path}`;
      if (ct.includes(`PartName="${partName}"`)) continue;
      let contentType: string;
      if (/pivotCacheDefinition\d+\.xml$/.test(path)) contentType = CT_PIVOT_CACHE_DEF;
      else if (/pivotCacheRecords\d+\.xml$/.test(path)) contentType = CT_PIVOT_CACHE_REC;
      else if (/pivotTable\d+\.xml$/.test(path)) contentType = CT_PIVOT_TABLE;
      else continue;
      inserts.push(`<Override PartName="${partName}" ContentType="${contentType}"/>`);
    }
    if (inserts.length > 0) {
      ct = ct.replace('</Types>', `${inserts.join('')}</Types>`);
      zip.file('[Content_Types].xml', ct);
    }
  }

  // 3. Patch xl/_rels/workbook.xml.rels — append the captured
  //    pivotCacheDefinition rels with fresh rIds. Track the
  //    origId → newId mapping for the <pivotCaches> patch in step 4.
  const workbookRelsPath = 'xl/_rels/workbook.xml.rels';
  const workbookRelsEntry = zip.file(workbookRelsPath);
  const cacheIdRemap = new Map<string, string>();
  if (workbookRelsEntry && payload.workbookCacheRels.length > 0) {
    let rels = await workbookRelsEntry.async('string');
    const used = new Set<number>();
    for (const m of rels.matchAll(/Id="rId(\d+)"/g)) used.add(Number(m[1]));
    const assign = nextRidIssuer(used);
    const additions: string[] = [];
    for (const r of payload.workbookCacheRels) {
      const newId = assign();
      cacheIdRemap.set(r.origId, newId);
      additions.push(
        `<Relationship Id="${newId}" Type="${REL_TYPE_PIVOT_CACHE_DEF}" Target="${r.target}"/>`,
      );
    }
    rels = rels.replace('</Relationships>', `${additions.join('')}</Relationships>`);
    zip.file(workbookRelsPath, rels);
  }

  // 4. Patch xl/workbook.xml — inject <pivotCaches> with the
  //    remapped rIds. The captured snippet has the original rIds;
  //    we substitute them in place via the cacheIdRemap.
  const workbookXmlEntry = zip.file('xl/workbook.xml');
  if (workbookXmlEntry && payload.workbookPivotCachesXml) {
    let workbookXml = await workbookXmlEntry.async('string');
    // Remap r:id references inside the captured snippet.
    let snippet = payload.workbookPivotCachesXml;
    for (const [origId, newId] of cacheIdRemap) {
      // Anchor on quote boundaries so we don't accidentally rewrite
      // a longer rId that prefix-overlaps (rId1 vs rId10).
      snippet = snippet.replaceAll(`"${origId}"`, `"${newId}"`);
    }
    // ExcelJS regenerates workbook.xml so it never contains a
    // <pivotCaches> element. Inject ours immediately after </sheets>
    // (the OOXML-correct position).
    if (!workbookXml.includes('<pivotCaches')) {
      workbookXml = workbookXml.replace('</sheets>', `</sheets>${snippet}`);
      zip.file('xl/workbook.xml', workbookXml);
    }
  }

  // 5. Patch each sheet's .rels — append captured pivotTable rels
  //    with fresh rIds. Sheet matching is by NAME via the regenerated
  //    workbook.xml; sheet file names may have shifted.
  if (workbookXmlEntry && payload.perSheet.length > 0) {
    const workbookXml = await workbookXmlEntry.async('string');
    const workbookRels = workbookRelsEntry
      ? parseRelationships(await zip.file(workbookRelsPath)!.async('string'))
      : [];

    const sheetByName = new Map<string, string>();
    for (const m of workbookXml.matchAll(/<sheet\b[^>]*\/>/g)) {
      const tag = m[0];
      const name = tag.match(/name="([^"]+)"/)?.[1];
      const rId =
        tag.match(/r:id="([^"]+)"/i)?.[1] ?? tag.match(/r:Id="([^"]+)"/)?.[1];
      if (!name || !rId) continue;
      const rel = workbookRels.find((r) => r.id === rId);
      if (!rel) continue;
      sheetByName.set(name, rel.target);
    }

    for (const sheet of payload.perSheet) {
      const targetPath = sheetByName.get(sheet.sheetName);
      if (!targetPath) continue; // sheet renamed or removed — skip
      const fileName = targetPath.split('/').pop()!;
      const sheetRelsPath = `xl/worksheets/_rels/${fileName}.rels`;
      let sheetRels =
        (await zip.file(sheetRelsPath)?.async('string')) ??
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
      const used = new Set<number>();
      for (const m of sheetRels.matchAll(/Id="rId(\d+)"/g)) used.add(Number(m[1]));
      const assign = nextRidIssuer(used);
      const additions: string[] = [];
      for (const r of sheet.pivotTableRels) {
        const newId = assign();
        additions.push(
          `<Relationship Id="${newId}" Type="${REL_TYPE_PIVOT_TABLE}" Target="${r.target}"/>`,
        );
      }
      if (sheetRels.includes('</Relationships>')) {
        sheetRels = sheetRels.replace(
          '</Relationships>',
          `${additions.join('')}</Relationships>`,
        );
      } else {
        sheetRels += `\n${additions.join('')}`;
      }
      zip.file(sheetRelsPath, sheetRels);
    }
  }

  // unused — suppress TS unused-var noise on the helper. (We use
  // extractRid + the constants in other branches of the file.)
  void extractRid;
}
