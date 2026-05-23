import JSZip from 'jszip';
import type { IWorkbookData } from '@univerjs/core';

/**
 * Sidecar resource that carries raw OOXML parts ExcelJS silently drops.
 * Today: xl/vbaProject.bin only (macros). Files round-trip as .xlsm so
 * the macros still run when re-opened in real Excel — we never execute
 * the VBA, just preserve the bytes. Complex pivot cache passthrough is
 * deferred to P6.1 (rel renumbering + workbook.xml surgery).
 */
export const XLSX_PASSTHROUGH_RESOURCE = '__casual_sheets_xlsx_passthrough__';

export type XlsxPassthroughPayload = {
  /** base64-encoded contents of xl/vbaProject.bin */
  vba?: { binBase64: string };
};

const VBA_REL_TYPE =
  'http://schemas.microsoft.com/office/2006/relationships/vbaProject';
const VBA_CONTENT_TYPE = 'application/vnd.ms-office.vbaProject';

export const XLSX_MIME =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
export const XLSM_MIME = 'application/vnd.ms-excel.sheet.macroEnabled.12';

export function mimeForPassthrough(
  payload: XlsxPassthroughPayload | undefined,
): string {
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
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(buffer);
  } catch {
    return undefined;
  }
  const vbaFile = zip.file('xl/vbaProject.bin');
  if (!vbaFile) return undefined;
  const binBase64 = await vbaFile.async('base64');
  return { vba: { binBase64 } };
}

export function mergePassthroughIntoResources(
  resources: IWorkbookData['resources'],
  payload: XlsxPassthroughPayload | undefined,
): IWorkbookData['resources'] {
  if (!payload) return resources;
  const filtered = (resources ?? []).filter(
    (r) => r.name !== XLSX_PASSTHROUGH_RESOURCE,
  );
  return [
    ...filtered,
    { name: XLSX_PASSTHROUGH_RESOURCE, data: JSON.stringify(payload) },
  ];
}

export function readPassthroughFromSnapshot(
  data: IWorkbookData,
): XlsxPassthroughPayload | undefined {
  const entry = data.resources?.find(
    (r) => r.name === XLSX_PASSTHROUGH_RESOURCE,
  );
  if (!entry?.data) return undefined;
  try {
    return JSON.parse(entry.data) as XlsxPassthroughPayload;
  } catch {
    return undefined;
  }
}

const REL_TYPE_REGEX_ESCAPE = (s: string) =>
  s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

/**
 * Re-inject captured OOXML parts into the ExcelJS-written buffer. The
 * input buffer is read but not mutated; a fresh ArrayBuffer is returned.
 * Pass-through is a no-op when payload is empty.
 */
export async function applyPassthroughToXlsxBuffer(
  excelJsBuffer: ArrayBuffer | Uint8Array,
  payload: XlsxPassthroughPayload | undefined,
): Promise<ArrayBuffer> {
  if (!payload?.vba) {
    if (excelJsBuffer instanceof ArrayBuffer) return excelJsBuffer;
    return excelJsBuffer.buffer.slice(
      excelJsBuffer.byteOffset,
      excelJsBuffer.byteOffset + excelJsBuffer.byteLength,
    ) as ArrayBuffer;
  }

  const zip = await JSZip.loadAsync(excelJsBuffer);

  zip.file('xl/vbaProject.bin', payload.vba.binBase64, { base64: true });

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

  // xl/_rels/workbook.xml.rels — append a vbaProject relationship with
  // the next-free rId. ExcelJS-written rels file already declares the
  // sheet / styles / theme / sharedStrings parts; we just need a unique
  // id that doesn't collide.
  const relsPath = 'xl/_rels/workbook.xml.rels';
  const relsEntry = zip.file(relsPath);
  if (relsEntry) {
    let rels = await relsEntry.async('string');
    const vbaRelTypeRegex = new RegExp(
      `Type="${REL_TYPE_REGEX_ESCAPE(VBA_REL_TYPE)}"`,
    );
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

  return zip.generateAsync({ type: 'arraybuffer' });
}
