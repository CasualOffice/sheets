import JSZip from 'jszip';
import type { IWorkbookData } from '@univerjs/core';
import {
  applyPivotsToZip,
  capturePivotsFromBuffer,
  type PivotPassthroughPayload,
} from './pivot-passthrough';
import { applyDataBarsToZip, type DataBarEntry } from './databar-passthrough';
import { applyDxfCfRulesToZip, type DxfCfRule } from './cf-dxf-passthrough';
import {
  applyDrawingsToZip,
  captureDrawingsFromBuffer,
  type DrawingPassthroughPayload,
} from './drawing-passthrough';
import {
  applyExternalLinksToZip,
  captureExternalLinksFromBuffer,
  type ExternalLinkPassthroughPayload,
} from './external-link-passthrough';

/**
 * Sidecar resource that carries raw OOXML parts ExcelJS silently drops.
 *
 * Today we passthrough:
 *  - `xl/vbaProject.bin` — `.xlsm` macros round-trip byte-equal.
 *  - `xl/pivotCaches/**` + `xl/pivotTables/**` — pivot definitions
 *    survive a round-trip so Excel still sees the file as having
 *    pivot tables (the materialised cells already survive via the
 *    normal cell pipeline; this re-instates the metadata Excel
 *    needs to render filter dropdowns, refresh, etc).
 *
 * We never execute VBA. The pivot OOXML is treated as opaque bytes —
 * Univer doesn't render pivots from the OOXML, only the cells; this
 * preserves the metadata for Excel's benefit.
 */
export const XLSX_PASSTHROUGH_RESOURCE = '__casual_sheets_xlsx_passthrough__';

export type XlsxPassthroughPayload = {
  /** base64-encoded contents of xl/vbaProject.bin */
  vba?: { binBase64: string };
  /** raw OOXML pivot machinery — see pivot-passthrough.ts */
  pivots?: PivotPassthroughPayload;
  /** data-bar CF blocks to splice in (ExcelJS can't write them) — keyed by
   *  sheet name; see databar-passthrough.ts */
  dataBars?: Record<string, DataBarEntry[]>;
  /** duplicate/unique CF blocks (+ their dxf styles) to splice in — keyed by
   *  sheet name; see cf-dxf-passthrough.ts */
  dxfCfRules?: Record<string, DxfCfRule[]>;
  /** embedded images / shapes (xl/media + xl/drawings) — Univer has no drawing
   *  model so they'd be dropped on save; see drawing-passthrough.ts */
  drawings?: DrawingPassthroughPayload;
  /** external-workbook links (xl/externalLinks) — ExcelJS drops them, breaking
   *  `[N]Sheet!A1` formulas; see external-link-passthrough.ts */
  externalLinks?: ExternalLinkPassthroughPayload;
};

const VBA_REL_TYPE = 'http://schemas.microsoft.com/office/2006/relationships/vbaProject';
const VBA_CONTENT_TYPE = 'application/vnd.ms-office.vbaProject';

export const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
export const XLSM_MIME = 'application/vnd.ms-excel.sheet.macroEnabled.12';

export function mimeForPassthrough(payload: XlsxPassthroughPayload | undefined): string {
  return payload?.vba ? XLSM_MIME : XLSX_MIME;
}

export function extensionForPassthrough(
  payload: XlsxPassthroughPayload | undefined,
): 'xlsx' | 'xlsm' {
  return payload?.vba ? 'xlsm' : 'xlsx';
}

export async function capturePassthroughFromBuffer(
  buffer: ArrayBuffer,
): Promise<XlsxPassthroughPayload | undefined> {
  // VBA — single binary part. Cheap probe; bail early if neither
  // VBA nor pivots are present.
  let vba: XlsxPassthroughPayload['vba'];
  try {
    const zip = await JSZip.loadAsync(buffer);
    const vbaFile = zip.file('xl/vbaProject.bin');
    if (vbaFile) {
      vba = { binBase64: await vbaFile.async('base64') };
    }
  } catch {
    return undefined;
  }

  // Pivots — re-opens the buffer internally. Could be optimised to a
  // single zip read but the second open is ~tens of ms even for big
  // files and the symmetry with VBA is more important.
  const pivots = await capturePivotsFromBuffer(buffer);
  const drawings = await captureDrawingsFromBuffer(buffer);
  const externalLinks = await captureExternalLinksFromBuffer(buffer);

  if (!vba && !pivots && !drawings && !externalLinks) return undefined;
  return { vba, pivots, drawings, externalLinks };
}

export function mergePassthroughIntoResources(
  resources: IWorkbookData['resources'],
  payload: XlsxPassthroughPayload | undefined,
): IWorkbookData['resources'] {
  if (!payload) return resources;
  const filtered = (resources ?? []).filter((r) => r.name !== XLSX_PASSTHROUGH_RESOURCE);
  return [...filtered, { name: XLSX_PASSTHROUGH_RESOURCE, data: JSON.stringify(payload) }];
}

export function readPassthroughFromSnapshot(
  data: IWorkbookData,
): XlsxPassthroughPayload | undefined {
  const entry = data.resources?.find((r) => r.name === XLSX_PASSTHROUGH_RESOURCE);
  if (!entry?.data) return undefined;
  try {
    return JSON.parse(entry.data) as XlsxPassthroughPayload;
  } catch {
    return undefined;
  }
}

const REL_TYPE_REGEX_ESCAPE = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

/**
 * Re-inject every captured OOXML payload into the ExcelJS-written
 * buffer in one pass. VBA + pivots share a single JSZip session so
 * the [Content_Types].xml + workbook.xml.rels patches compose cleanly
 * (each patch sees the previous step's writes).
 *
 * The input buffer is read but not mutated; a fresh ArrayBuffer is
 * returned. No-op when the payload is empty.
 */
export async function applyPassthroughToXlsxBuffer(
  excelJsBuffer: ArrayBuffer | Uint8Array,
  payload: XlsxPassthroughPayload | undefined,
): Promise<ArrayBuffer> {
  const hasDataBars = payload?.dataBars && Object.keys(payload.dataBars).length > 0;
  const hasDxfCf = payload?.dxfCfRules && Object.keys(payload.dxfCfRules).length > 0;
  const hasDrawings = payload?.drawings && Object.keys(payload.drawings.parts).length > 0;
  const hasExternalLinks =
    payload?.externalLinks && Object.keys(payload.externalLinks.parts).length > 0;
  if (
    !payload?.vba &&
    !payload?.pivots &&
    !hasDataBars &&
    !hasDxfCf &&
    !hasDrawings &&
    !hasExternalLinks
  ) {
    if (excelJsBuffer instanceof ArrayBuffer) return excelJsBuffer;
    return excelJsBuffer.buffer.slice(
      excelJsBuffer.byteOffset,
      excelJsBuffer.byteOffset + excelJsBuffer.byteLength,
    ) as ArrayBuffer;
  }

  const zip = await JSZip.loadAsync(excelJsBuffer);

  if (payload.vba) await applyVbaToZip(zip, payload.vba);
  if (payload.pivots) await applyPivotsToZip(zip, payload.pivots);
  if (payload.dataBars) await applyDataBarsToZip(zip, payload.dataBars);
  if (payload.dxfCfRules) await applyDxfCfRulesToZip(zip, payload.dxfCfRules);
  if (payload.drawings) await applyDrawingsToZip(zip, payload.drawings);
  if (payload.externalLinks) await applyExternalLinksToZip(zip, payload.externalLinks);

  return zip.generateAsync({ type: 'arraybuffer' });
}

async function applyVbaToZip(
  zip: JSZip,
  vba: NonNullable<XlsxPassthroughPayload['vba']>,
): Promise<void> {
  zip.file('xl/vbaProject.bin', vba.binBase64, { base64: true });

  // [Content_Types].xml — add Override for /xl/vbaProject.bin if missing.
  const ctEntry = zip.file('[Content_Types].xml');
  if (ctEntry) {
    let ct = await ctEntry.async('string');
    if (!/PartName="\/xl\/vbaProject\.bin"/i.test(ct)) {
      const override = `<Override PartName="/xl/vbaProject.bin" ContentType="${VBA_CONTENT_TYPE}"/>`;
      ct = ct.replace('</Types>', `${override}</Types>`);
      zip.file('[Content_Types].xml', ct);
    }
  }

  // xl/_rels/workbook.xml.rels — append a vbaProject relationship
  // with the next-free rId. ExcelJS-written rels already declares
  // sheet / styles / theme / sharedStrings; just need a unique id.
  const relsPath = 'xl/_rels/workbook.xml.rels';
  const relsEntry = zip.file(relsPath);
  if (relsEntry) {
    let rels = await relsEntry.async('string');
    const vbaRelTypeRegex = new RegExp(`Type="${REL_TYPE_REGEX_ESCAPE(VBA_REL_TYPE)}"`);
    if (!vbaRelTypeRegex.test(rels)) {
      const used = new Set<number>();
      for (const m of rels.matchAll(/Id="rId(\d+)"/g)) used.add(Number(m[1]));
      let next = 1;
      while (used.has(next)) next++;
      const rel = `<Relationship Id="rId${next}" Type="${VBA_REL_TYPE}" Target="vbaProject.bin"/>`;
      rels = rels.replace('</Relationships>', `${rel}</Relationships>`);
      zip.file(relsPath, rels);
    }
  }
}
