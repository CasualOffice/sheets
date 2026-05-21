import { CustomRangeType, type IWorkbookData } from '@univerjs/core';
import type { FUniver } from '@univerjs/core/facade';
import { ISheetClipboardService } from '@univerjs/sheets-ui';
import { SheetTableService } from '@univerjs/sheets-table';

type HyperLinkDump = {
  subUnitId: string;
  row: number;
  column: number;
  payload: string;
  display?: string;
};

type ParsedClipboardCell = {
  row: number;
  column: number;
  v?: unknown;
  s?: unknown;
  f?: unknown;
  rowSpan?: number;
  colSpan?: number;
};

type ParsedClipboardDump = {
  cells: ParsedClipboardCell[];
  rowProperties: Array<Record<string, unknown>>;
  colProperties: Array<Record<string, unknown>>;
};

declare global {
  interface Window {
    __univerAPI?: FUniver;
    __getTableStyleId__?: (tableId: string) => string | undefined;
    __getHyperLinks__?: () => HyperLinkDump[];
    __legacyPasteHtml__?: (html: string, text?: string) => Promise<boolean>;
    __parseHtmlClipboard__?: (html: string) => ParsedClipboardDump | null;
  }
}

/**
 * DEV-only window helpers used by e2e specs. Anything that needs to reach
 * into Univer's internals from a Playwright test belongs here, not in
 * production code paths.
 *
 * Currently:
 *   - __univerAPI exposes the FUniver facade.
 *   - __getTableStyleId__ exposes the underlying Table's tableStyleId, which
 *     FWorkbook.getTableList intentionally strips from its public projection.
 */
export function installDevHelpers(api: FUniver): () => void {
  if (!import.meta.env.DEV) return () => {};

  window.__univerAPI = api;
  window.__getTableStyleId__ = (tableId) => {
    const wb = api.getActiveWorkbook();
    if (!wb) return undefined;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const svc = (wb as any)._injector?.get(SheetTableService) as
      | {
          _tableManager?: {
            getTable: (u: string, t: string) =>
              | { getTableStyleId: () => string }
              | undefined;
          };
        }
      | undefined;
    return svc?._tableManager?.getTable(wb.getId(), tableId)?.getTableStyleId();
  };
  // Dumps every hyperlink across the active workbook. Hyperlinks live in
  // `cell.p.body.customRanges` (the rich-text custom range model), not in
  // HyperLinkModel — AddHyperLinkCommand writes the cell body and skips the
  // model, so the model is unreliable as a source-of-truth.
  window.__getHyperLinks__ = () => {
    const wb = api.getActiveWorkbook();
    if (!wb) return [];
    const snap = wb.save() as IWorkbookData;
    const out: HyperLinkDump[] = [];
    for (const sheetId of snap.sheetOrder ?? []) {
      const wsd = snap.sheets?.[sheetId];
      if (!wsd?.cellData) continue;
      const cellData = wsd.cellData as Record<
        string,
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        Record<string, { p?: any }>
      >;
      for (const rKey of Object.keys(cellData)) {
        const r = Number(rKey);
        const row = cellData[rKey];
        for (const cKey of Object.keys(row)) {
          const c = Number(cKey);
          const body = row[cKey]?.p?.body;
          const ranges: Array<{
            startIndex: number;
            endIndex: number;
            rangeType: CustomRangeType;
            properties?: { url?: string };
          }> = body?.customRanges ?? [];
          for (const cr of ranges) {
            if (cr.rangeType !== CustomRangeType.HYPERLINK) continue;
            const url = cr.properties?.url;
            if (typeof url !== 'string' || !url) continue;
            const dataStream: string = body?.dataStream ?? '';
            out.push({
              subUnitId: sheetId,
              row: r,
              column: c,
              payload: url,
              display: dataStream.slice(cr.startIndex, cr.endIndex + 1),
            });
          }
        }
      }
    }
    return out;
  };
  window.__legacyPasteHtml__ = async (html, text = '') => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const injector = (api as any)._injector as
      | { get: (token: unknown) => unknown }
      | undefined;
    if (!injector) return false;
    const svc = injector.get(ISheetClipboardService) as
      | {
          legacyPaste: (html: string, text?: string, files?: File[]) => Promise<boolean>;
        }
      | undefined;
    if (!svc?.legacyPaste) return false;
    return svc.legacyPaste(html, text, []);
  };
  window.__parseHtmlClipboard__ = (html) => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const injector = (api as any)._injector as
      | { get: (token: unknown) => unknown }
      | undefined;
    if (!injector) return null;
    const svc = injector.get(ISheetClipboardService) as
      | {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          _htmlToUSM?: { convert: (html: string) => any };
        }
      | undefined;
    const parsed = svc?._htmlToUSM?.convert(html);
    if (!parsed?.cellMatrix?.forValue) return null;
    const cells: ParsedClipboardCell[] = [];
    parsed.cellMatrix.forValue(
      (
        row: number,
        column: number,
        cell: {
          v?: unknown;
          s?: unknown;
          f?: unknown;
          rowSpan?: number;
          colSpan?: number;
        } | null,
      ) => {
        if (!cell) return;
        cells.push({
          row,
          column,
          v: cell.v,
          s: cell.s,
          f: cell.f,
          rowSpan: cell.rowSpan,
          colSpan: cell.colSpan,
        });
      },
    );
    return {
      cells,
      rowProperties: parsed.rowProperties ?? [],
      colProperties: parsed.colProperties ?? [],
    };
  };

  return () => {
    delete window.__univerAPI;
    delete window.__getTableStyleId__;
    delete window.__getHyperLinks__;
    delete window.__legacyPasteHtml__;
    delete window.__parseHtmlClipboard__;
  };
}
