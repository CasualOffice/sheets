import { CustomRangeType, type IWorkbookData } from '@univerjs/core';
import type { FUniver } from '@univerjs/core/facade';
import { SheetTableService } from '@univerjs/sheets-table';

type HyperLinkDump = {
  subUnitId: string;
  row: number;
  column: number;
  payload: string;
  display?: string;
};

declare global {
  interface Window {
    __univerAPI?: FUniver;
    __getTableStyleId__?: (tableId: string) => string | undefined;
    __getHyperLinks__?: () => HyperLinkDump[];
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

  return () => {
    delete window.__univerAPI;
    delete window.__getTableStyleId__;
    delete window.__getHyperLinks__;
  };
}
