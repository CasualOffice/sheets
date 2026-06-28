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

import { CustomRangeType, IMentionIOService, type IWorkbookData } from '@univerjs/core';
import type { FUniver } from '@univerjs/core/facade';
import { ISheetClipboardService } from '@univerjs/sheets-ui';
import { SheetTableService } from '@univerjs/sheets-table';
import { ConditionalFormattingService } from '@univerjs/sheets-conditional-formatting';
import { ensurePluginByName, type LazyPluginGroup } from '@casualoffice/sheets/univer';
import {
  setMentionProvider,
  filterMentionCandidates,
  type MentionCandidate,
} from '@casualoffice/sheets/sheets';

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
    __ensurePlugin__?: (group: LazyPluginGroup) => Promise<void>;
    __getTableStyleId__?: (tableId: string) => string | undefined;
    __getHyperLinks__?: () => HyperLinkDump[];
    __legacyPasteHtml__?: (html: string, text?: string) => Promise<boolean>;
    __parseHtmlClipboard__?: (html: string) => ParsedClipboardDump | null;
    /** Test hook: install a fixed @mention candidate list. */
    __setMentionProvider__?: (candidates: MentionCandidate[]) => void;
    /** Test hook: resolve the live IMentionIOService and return candidate labels. */
    __mentionList__?: (search?: string) => Promise<string[] | null>;
    /** Test hook: compose the conditional-formatting style for a cell on the
     *  active sheet — proves an imported CF rule actually painted (the highlight
     *  is canvas-drawn, not in cell data). Returns `null` when no rule matches. */
    __composeCfStyle__?: (row: number, col: number) => unknown;
  }
}

/**
 * Window helpers used by e2e specs AND prod debugging. Anything that
 * needs to reach into Univer's internals from a Playwright test belongs
 * here, not in production code paths.
 *
 * `__univerAPI` is exposed in BOTH dev and prod — it's the FUniver
 * facade with no secrets, and having it available in the deployed
 * docker build is what lets us run regression tests against the
 * actual prod bundle. The rest of the helpers stay DEV-only because
 * they expose internal injector tokens that shouldn't be reachable
 * from random page scripts in production.
 */
export function installDevHelpers(api: FUniver): () => void {
  // Always expose the facade so prod-build regression tests can reach in.
  window.__univerAPI = api;
  // Expose the lazy-plugin loader in BOTH dev and prod so e2e specs can
  // deterministically wait for a feature plugin (drawings, CF, DV, etc.)
  // to be registered before invoking its command. Without this, tests
  // race `idleLoadAll`, which is fine 95% of the time but flakes the
  // first run after a cold boot.
  window.__ensurePlugin__ = ensurePluginByName;
  if (!import.meta.env.DEV) {
    return () => {
      delete window.__univerAPI;
      delete window.__ensurePlugin__;
    };
  }
  window.__getTableStyleId__ = (tableId) => {
    const wb = api.getActiveWorkbook();
    if (!wb) return undefined;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const svc = (wb as any)._injector?.get(SheetTableService) as
      | {
          _tableManager?: {
            getTable: (u: string, t: string) => { getTableStyleId: () => string } | undefined;
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
    const injector = (api as any)._injector as { get: (token: unknown) => unknown } | undefined;
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
    const injector = (api as any)._injector as { get: (token: unknown) => unknown } | undefined;
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
  // @mention test hooks. `__setMentionProvider__` installs a fixed candidate
  // list (the app's real provider only has data inside a room); `__mentionList__`
  // resolves the LIVE IMentionIOService off the injector and returns the labels
  // it produces — proving the CasualSheets DI override is in effect and reads
  // our provider, end-to-end through Univer, without driving the in-cell popup.
  window.__setMentionProvider__ = (candidates) => {
    setMentionProvider((search) => filterMentionCandidates(candidates, search));
  };
  window.__composeCfStyle__ = (row, col) => {
    const wb = api.getActiveWorkbook();
    if (!wb) return null;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const injector = (api as any)._injector as { get: (token: unknown) => unknown } | undefined;
    if (!injector) return null;
    const svc = injector.get(ConditionalFormattingService) as
      | { composeStyle: (u: string, s: string, r: number, c: number) => unknown }
      | undefined;
    const ws = wb.getActiveSheet();
    if (!svc?.composeStyle || !ws) return null;
    return svc.composeStyle(wb.getId(), ws.getSheetId(), row, col);
  };
  window.__mentionList__ = async (search = '') => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const injector = (api as any)._injector;
    const svc = injector?.get?.(IMentionIOService) as
      | {
          list: (p: {
            search?: string;
          }) => Promise<{ list: Array<{ mentions: Array<{ label: string }> }> }>;
        }
      | undefined;
    if (!svc?.list) return null;
    const res = await svc.list({ search });
    return res.list.flatMap((g) => g.mentions.map((m) => m.label));
  };

  return () => {
    delete window.__univerAPI;
    delete window.__ensurePlugin__;
    delete window.__getTableStyleId__;
    delete window.__getHyperLinks__;
    delete window.__legacyPasteHtml__;
    delete window.__parseHtmlClipboard__;
    delete window.__setMentionProvider__;
    delete window.__mentionList__;
    delete window.__composeCfStyle__;
  };
}
