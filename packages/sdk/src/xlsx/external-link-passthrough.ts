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
 * External-workbook-link passthrough.
 *
 * A formula like `=[1]Sheet1!A1` references another workbook. The `[1]` is a
 * 1-based index into `<externalReferences>` in `xl/workbook.xml`, which points
 * (via `xl/_rels/workbook.xml.rels`) at `xl/externalLinks/externalLink{N}.xml`
 * — the part that records the source workbook path + the cached cell values.
 *
 * ExcelJS has no external-link model: it rebuilds the workbook without
 * `xl/externalLinks/**` and without `<externalReferences>`, so on save every
 * such formula loses its target and resolves to `#REF!`. The formula *text*
 * survives (Univer keeps it), but the link it indexes is gone — silent data
 * loss / corruption.
 *
 * This captures `xl/externalLinks/**` verbatim at parse time, in the original
 * reference order, and re-injects them at export: restores the parts, patches
 * `[Content_Types].xml`, re-creates the workbook→externalLink relationships,
 * and rebuilds `<externalReferences>` in the same order so the `[N]` indices
 * still resolve. Workbook-level (mirrors pivot/drawing passthrough); the parts
 * are opaque bytes — we don't refresh the cached values.
 */

const REL_TYPE_EXTERNAL_LINK =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink';
const CT_EXTERNAL_LINK =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml';

export interface ExternalLinkPassthroughPayload {
  /** xl/externalLinks/** parts (the link xml + its _rels), keyed by zip path; base64. */
  parts: Record<string, string>;
  /** Ordered list of the externalLink part paths, matching the original
   *  `<externalReferences>` order — preserves the `[N]` formula indices. */
  order: string[];
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

/** Resolve a workbook-relative rel Target (e.g. `externalLinks/externalLink1.xml`
 *  or `/xl/externalLinks/externalLink1.xml`) to an absolute zip path. */
function resolveFromXl(target: string): string {
  if (target.startsWith('/')) return target.replace(/^\//, '');
  const stack = 'xl'.split('/');
  for (const seg of target.split('/')) {
    if (seg === '..') stack.pop();
    else if (seg !== '.') stack.push(seg);
  }
  return stack.join('/');
}

/**
 * Capture every external-link part + their reference order from a source xlsx.
 * Returns undefined when the workbook has no external links.
 */
export async function captureExternalLinksFromBuffer(
  buffer: ArrayBuffer,
): Promise<ExternalLinkPassthroughPayload | undefined> {
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(buffer);
  } catch {
    return undefined;
  }

  const parts: Record<string, string> = {};
  for (const path of Object.keys(zip.files)) {
    if (zip.files[path].dir) continue;
    if (/^xl\/externalLinks\//.test(path)) {
      parts[path] = await zip.files[path].async('base64');
    }
  }
  if (Object.keys(parts).length === 0) return undefined;

  // Reconstruct the reference order from <externalReferences> → rId → target.
  // Fall back to numeric filename order if the workbook block can't be read.
  const order: string[] = [];
  const workbookXml = await zip.file('xl/workbook.xml')?.async('string');
  const relsXml = await zip.file('xl/_rels/workbook.xml.rels')?.async('string');
  if (workbookXml && relsXml) {
    const rels = parseRels(relsXml);
    const refsBlock = workbookXml.match(
      /<externalReferences>([\s\S]*?)<\/externalReferences>/,
    )?.[1];
    if (refsBlock) {
      for (const m of refsBlock.matchAll(/<externalReference\b[^>]*r:id="([^"]+)"/gi)) {
        const target = rels.find((r) => r.id === m[1])?.target;
        if (target) order.push(resolveFromXl(target));
      }
    }
  }
  if (order.length === 0) {
    // Fall back: numeric order of the main link parts.
    order.push(
      ...Object.keys(parts)
        .filter((p) => /^xl\/externalLinks\/externalLink\d+\.xml$/.test(p))
        .sort((a, b) => Number(a.match(/(\d+)\.xml$/)![1]) - Number(b.match(/(\d+)\.xml$/)![1])),
    );
  }

  return { parts, order };
}

// Workbook child elements that must follow <externalReferences> in CT_Workbook
// order; insert the block just before the first one present.
const AFTER_EXTERNAL_REFS = [
  '<definedNames',
  '<calcPr',
  '<oleSize',
  '<customWorkbookViews',
  '<pivotCaches',
  '<smartTagPr',
  '<smartTagTypes',
  '<webPublishing',
  '<fileRecoveryPr',
  '<webPublishObjects',
  '<extLst',
  '</workbook>',
];

/**
 * Re-inject captured external-link parts into the ExcelJS-written zip and
 * rebuild `<externalReferences>`. Mutates `zip` in place (composes with the
 * other passthrough steps). No-op if the workbook already declares external
 * references.
 */
export async function applyExternalLinksToZip(
  zip: JSZip,
  payload: ExternalLinkPassthroughPayload,
): Promise<void> {
  // 1. Restore every external-link part verbatim.
  for (const [path, base64] of Object.entries(payload.parts)) {
    zip.file(path, base64, { base64: true });
  }

  // 2. [Content_Types].xml — Override per externalLink{N}.xml (not the _rels).
  const ctEntry = zip.file('[Content_Types].xml');
  if (ctEntry) {
    let ct = await ctEntry.async('string');
    const inserts: string[] = [];
    for (const path of Object.keys(payload.parts)) {
      if (!/^xl\/externalLinks\/externalLink\d+\.xml$/.test(path)) continue;
      const partName = `/${path}`;
      if (ct.includes(`PartName="${partName}"`)) continue;
      inserts.push(`<Override PartName="${partName}" ContentType="${CT_EXTERNAL_LINK}"/>`);
    }
    if (inserts.length > 0) {
      ct = ct.replace('</Types>', `${inserts.join('')}</Types>`);
      zip.file('[Content_Types].xml', ct);
    }
  }

  // 3. workbook.xml.rels — one relationship per ordered link, fresh rIds.
  const relsPath = 'xl/_rels/workbook.xml.rels';
  let relsXml =
    (await zip.file(relsPath)?.async('string')) ??
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  const used = new Set<number>();
  for (const m of relsXml.matchAll(/Id="rId(\d+)"/g)) used.add(Number(m[1]));
  let next = 1;
  const freshRid = (): string => {
    while (used.has(next)) next++;
    used.add(next);
    return `rId${next}`;
  };

  const orderedRids: string[] = [];
  const relInserts: string[] = [];
  for (const partPath of payload.order) {
    const rId = freshRid();
    orderedRids.push(rId);
    // Target is relative to xl/ (workbook.xml.rels base).
    const target = partPath.replace(/^xl\//, '');
    relInserts.push(
      `<Relationship Id="${rId}" Type="${REL_TYPE_EXTERNAL_LINK}" Target="${target}"/>`,
    );
  }
  relsXml = relsXml.replace('</Relationships>', `${relInserts.join('')}</Relationships>`);
  zip.file(relsPath, relsXml);

  // 4. workbook.xml — inject <externalReferences> in schema order.
  const wbEntry = zip.file('xl/workbook.xml');
  if (wbEntry) {
    let wb = await wbEntry.async('string');
    if (!wb.includes('<externalReferences>') && orderedRids.length > 0) {
      const block = `<externalReferences>${orderedRids
        .map((rId) => `<externalReference r:id="${rId}"/>`)
        .join('')}</externalReferences>`;
      const at = AFTER_EXTERNAL_REFS.map((t) => wb.indexOf(t)).filter((i) => i !== -1);
      const idx = at.length ? Math.min(...at) : -1;
      wb = idx !== -1 ? wb.slice(0, idx) + block + wb.slice(idx) : wb;
      zip.file('xl/workbook.xml', wb);
    }
  }
}
