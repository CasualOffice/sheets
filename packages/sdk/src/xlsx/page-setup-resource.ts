import type ExcelJS from 'exceljs';
import type { IWorkbookData } from '@univerjs/core';

/**
 * Passthrough preservation for xlsx-native page-setup chrome —
 * orientation, paper size, margins, header/footer text. Univer
 * doesn't model any of this; the print dialog uses a per-user
 * localStorage slot. Without a passthrough, opening an xlsx that
 * was authored to print landscape and saving it again silently
 * resets to portrait.
 *
 * The round-trip stays inside our own resource sidecar
 * (`__casual_sheets_page_setup__`) — the data never reaches Univer
 * for editing; the parser stashes it on `IWorkbookData.resources`
 * and the exporter reads it back and applies to ExcelJS's native
 * `worksheet.pageSetup` + `worksheet.headerFooter`.
 *
 * If/when we add an in-app Page Setup editor, this resource is the
 * authoritative store — the localStorage fallback is per-user and
 * can't survive Save → Open in another tab.
 */

export const PAGE_SETUP_RESOURCE = '__casual_sheets_page_setup__';

export type SheetPageSetupV1 = {
  orientation?: 'landscape' | 'portrait';
  paperSize?: number;
  fitToPage?: boolean;
  fitToWidth?: number;
  fitToHeight?: number;
  scale?: number;
  printArea?: string;
  margins?: {
    top?: number;
    bottom?: number;
    left?: number;
    right?: number;
    header?: number;
    footer?: number;
  };
  headerFooter?: {
    oddHeader?: string;
    oddFooter?: string;
    evenHeader?: string;
    evenFooter?: string;
    differentFirst?: boolean;
    differentOddEven?: boolean;
  };
};

type Payload = {
  v: 1;
  sheets: Record<string, SheetPageSetupV1>;
};

function isLandscape(o: unknown): o is 'landscape' {
  return o === 'landscape';
}

// ExcelJS materialises a default page-setup + margins object on every
// worksheet during load, even when the source xlsx never authored one.
// Capturing those defaults would write a `__casual_sheets_page_setup__`
// resource to every workbook that opens here — visible noise in
// snapshot diffs and a failing assertion in `xlsx-hyperlinks.spec.ts`'s
// resources round-trip probe. Drop them so the sidecar only appears
// when the user (or the source xlsx) authored a non-default value.
const PAGE_SETUP_DEFAULTS = {
  orientation: 'portrait' as const,
  fitToPage: false,
  fitToWidth: 1,
  fitToHeight: 1,
  scale: 100,
};
const MARGIN_DEFAULTS = {
  top: 0.75,
  bottom: 0.75,
  left: 0.7,
  right: 0.7,
  header: 0.3,
  footer: 0.3,
};

/**
 * Walk the xlsx worksheets and pull each one's page-setup + header
 * /footer into a sidecar payload. Empty entries are dropped so the
 * resource isn't written for files that didn't carry any chrome.
 */
export function readPageSetupFromXlsx(
  wb: ExcelJS.Workbook,
  sheetIdForExcel: (excelId: number) => string,
): Record<string, SheetPageSetupV1> {
  const out: Record<string, SheetPageSetupV1> = {};
  for (const ws of wb.worksheets) {
    const ps = (ws as unknown as { pageSetup?: Record<string, unknown> }).pageSetup;
    const hf = (ws as unknown as { headerFooter?: Record<string, unknown> }).headerFooter;
    const entry: SheetPageSetupV1 = {};

    if (ps && typeof ps === 'object') {
      // Only `landscape` matters — `portrait` is the xlsx default and
      // would otherwise tag every workbook with a sidecar.
      if (isLandscape(ps.orientation)) entry.orientation = 'landscape';
      if (typeof ps.paperSize === 'number') entry.paperSize = ps.paperSize;
      if (ps.fitToPage === true) entry.fitToPage = true;
      if (typeof ps.fitToWidth === 'number' && ps.fitToWidth !== PAGE_SETUP_DEFAULTS.fitToWidth) {
        entry.fitToWidth = ps.fitToWidth as number;
      }
      if (typeof ps.fitToHeight === 'number' && ps.fitToHeight !== PAGE_SETUP_DEFAULTS.fitToHeight) {
        entry.fitToHeight = ps.fitToHeight as number;
      }
      if (typeof ps.scale === 'number' && ps.scale !== PAGE_SETUP_DEFAULTS.scale) {
        entry.scale = ps.scale as number;
      }
      if (typeof ps.printArea === 'string') entry.printArea = ps.printArea as string;
      const margins = ps.margins as Record<string, unknown> | undefined;
      if (margins && typeof margins === 'object') {
        const m: SheetPageSetupV1['margins'] = {};
        for (const k of ['top', 'bottom', 'left', 'right', 'header', 'footer'] as const) {
          const v = margins[k];
          if (typeof v === 'number' && v !== MARGIN_DEFAULTS[k]) m[k] = v;
        }
        if (Object.keys(m).length > 0) entry.margins = m;
      }
    }

    if (hf && typeof hf === 'object') {
      const headerFooter: SheetPageSetupV1['headerFooter'] = {};
      for (const k of ['oddHeader', 'oddFooter', 'evenHeader', 'evenFooter'] as const) {
        const v = hf[k];
        if (typeof v === 'string' && v.length > 0) headerFooter[k] = v as string;
      }
      if (hf.differentFirst === true) headerFooter.differentFirst = true;
      if (hf.differentOddEven === true) headerFooter.differentOddEven = true;
      if (Object.keys(headerFooter).length > 0) entry.headerFooter = headerFooter;
    }

    if (Object.keys(entry).length > 0) {
      out[sheetIdForExcel(ws.id)] = entry;
    }
  }
  return out;
}

/** Merge a synthesised page-setup map into the snapshot resources. */
export function mergePageSetupIntoResources(
  resources: IWorkbookData['resources'],
  payload: Record<string, SheetPageSetupV1>,
): IWorkbookData['resources'] {
  if (Object.keys(payload).length === 0) return resources;
  const existing = resources?.find((r) => r.name === PAGE_SETUP_RESOURCE);
  if (existing) return resources;
  const body: Payload = { v: 1, sheets: payload };
  const next = [...(resources ?? [])];
  next.push({ name: PAGE_SETUP_RESOURCE, data: JSON.stringify(body) });
  return next;
}

/** Read the page-setup resource off a snapshot. Tolerant of older /
 *  missing / malformed payloads. */
export function readPageSetupFromSnapshot(
  data: IWorkbookData,
): Record<string, SheetPageSetupV1> {
  const entry = data.resources?.find((r) => r.name === PAGE_SETUP_RESOURCE);
  if (!entry?.data) return {};
  try {
    const parsed = JSON.parse(entry.data) as Partial<Payload>;
    if (parsed?.v !== 1 || !parsed.sheets) return {};
    return parsed.sheets;
  } catch {
    return {};
  }
}

/** Apply a per-sheet entry onto an ExcelJS worksheet. */
export function applyPageSetupToXlsxWorksheet(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  ws: any,
  entry: SheetPageSetupV1,
): void {
  if (!entry) return;
  ws.pageSetup = ws.pageSetup ?? {};
  if (entry.orientation) ws.pageSetup.orientation = entry.orientation;
  if (typeof entry.paperSize === 'number') ws.pageSetup.paperSize = entry.paperSize;
  if (typeof entry.fitToPage === 'boolean') ws.pageSetup.fitToPage = entry.fitToPage;
  if (typeof entry.fitToWidth === 'number') ws.pageSetup.fitToWidth = entry.fitToWidth;
  if (typeof entry.fitToHeight === 'number') ws.pageSetup.fitToHeight = entry.fitToHeight;
  if (typeof entry.scale === 'number') ws.pageSetup.scale = entry.scale;
  if (typeof entry.printArea === 'string') ws.pageSetup.printArea = entry.printArea;
  if (entry.margins) {
    ws.pageSetup.margins = { ...(ws.pageSetup.margins ?? {}), ...entry.margins };
  }

  if (entry.headerFooter) {
    ws.headerFooter = ws.headerFooter ?? {};
    for (const k of ['oddHeader', 'oddFooter', 'evenHeader', 'evenFooter'] as const) {
      const v = entry.headerFooter[k];
      if (typeof v === 'string') ws.headerFooter[k] = v;
    }
    if (typeof entry.headerFooter.differentFirst === 'boolean') {
      ws.headerFooter.differentFirst = entry.headerFooter.differentFirst;
    }
    if (typeof entry.headerFooter.differentOddEven === 'boolean') {
      ws.headerFooter.differentOddEven = entry.headerFooter.differentOddEven;
    }
  }
}
