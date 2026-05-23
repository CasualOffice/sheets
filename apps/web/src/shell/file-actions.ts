import { CustomRangeType, type IWorkbookData } from '@univerjs/core';
import type { FUniver } from '@univerjs/core/facade';
import { IMessageService } from '@univerjs/ui';
import { MessageType } from '@univerjs/design';
import { workbookDataToXlsx, xlsxToWorkbookData } from '../xlsx';
import { timeIt } from '../perf';
import type { ExportExtras } from '../xlsx/export';
import type { OutlineState } from '../outline/types';
import type { ChartModel } from '../charts/types';
import { renderChartToPng, pixelsForChart } from '../charts/render-to-png';
import { recordRecentFile } from '../recent-files/store';
import type { PivotModel } from '../pivots/types';
import { discardAutosaveAfterExplicitSave } from '../autosave/useAutosave';
import {
  csvToWorkbookData,
  odsToWorkbookData,
  tsvToWorkbookData,
  workbookDataToDelimited,
  workbookDataToOds,
} from '../ods';

/**
 * File-level imperative actions. Pure functions — the caller owns React state
 * (e.g. lifting the workbook snapshot so a new Open replaces the active unit).
 */

/** Optional progress reporter called as the open flow transitions
 *  between phases (read buffer → parse → mount). Used by the loading
 *  overlay; ignored by callers that don't want UI feedback. */
export type OpenProgress = (phase: 'reading' | 'parsing' | 'mounting') => void;

/**
 * Hard limit on uploads. Past this size, Chrome reliably OOMs the tab
 * mid-parse — better to fail fast with a clear error than crash. The
 * supported ceiling matches the pipeline doc; soft warning lives in
 * the overlay (the user has time to read it during the parse stage).
 *
 * Both knobs are baked at Vite build time from `VITE_MAX_OPEN_MB` /
 * `VITE_SOFT_WARN_MB` so a self-host can raise them per their tab's
 * memory ceiling (or lower them on weaker hardware). Defaults are
 * sized for a typical desktop Chrome — pipeline doc §"Recommended
 * limits".
 *
 * If you raise these, also bump the server's multipart limit
 * (`MAX_UPLOAD_MB` env in `apps/server/src/index.ts`) so the co-edit
 * seed upload doesn't 413.
 */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const envMaxOpenMb = Number((import.meta.env as any).VITE_MAX_OPEN_MB);
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const envSoftWarnMb = Number((import.meta.env as any).VITE_SOFT_WARN_MB);
export const MAX_OPEN_BYTES =
  (Number.isFinite(envMaxOpenMb) && envMaxOpenMb > 0 ? envMaxOpenMb : 100) * 1024 * 1024;
export const SOFT_WARN_BYTES =
  (Number.isFinite(envSoftWarnMb) && envSoftWarnMb > 0 ? envSoftWarnMb : 25) * 1024 * 1024;

/**
 * Open a spreadsheet from disk. We auto-detect by file extension and
 * dispatch to the right parser. xlsx files go through ExcelJS (in a
 * worker); ods files through SheetJS Community.
 *
 * Throws if `file.size` exceeds `MAX_OPEN_BYTES` — the worker can't
 * hold the parsed `IWorkbookData` for a 100+ MB xlsx without OOM-ing
 * the tab. Caller is expected to catch and surface the error via the
 * loading overlay.
 */
export async function openSpreadsheetFile(
  file: File,
  onProgress?: OpenProgress,
): Promise<IWorkbookData> {
  console.info('[open] reading file', { name: file.name, size: file.size });
  if (file.size > MAX_OPEN_BYTES) {
    const mb = (file.size / (1024 * 1024)).toFixed(0);
    const cap = (MAX_OPEN_BYTES / (1024 * 1024)).toFixed(0);
    throw new Error(
      `File too large to open in the browser (${mb} MB). The current ceiling is ${cap} MB — past that, the worker can't allocate the parsed workbook without crashing the tab. Split the workbook or strip unused sheets before opening.`,
    );
  }
  onProgress?.('reading');
  const buf = await file.arrayBuffer();
  console.info('[open] buffer read', buf.byteLength, 'bytes — parsing');
  onProgress?.('parsing');
  const lower = file.name.toLowerCase();
  let data: IWorkbookData;
  if (lower.endsWith('.ods')) data = await odsToWorkbookData(buf);
  else if (lower.endsWith('.csv')) data = await csvToWorkbookData(buf);
  else if (lower.endsWith('.tsv') || lower.endsWith('.tab')) data = await tsvToWorkbookData(buf);
  else data = await xlsxToWorkbookData(buf);
  console.info('[open] parsed', { id: data.id, sheets: Object.keys(data.sheets ?? {}).length });
  data.name = file.name.replace(/\.(xlsx|ods|csv|tsv|tab)$/i, '');
  return data;
}

/** Back-compat alias — older callers reference openXlsx by name. */
export const openXlsx = openSpreadsheetFile;

export type WorkbookFormat = 'xlsx' | 'ods' | 'csv' | 'tsv';

function inferFormat(filename: string): WorkbookFormat {
  const lower = filename.toLowerCase();
  if (lower.endsWith('.ods')) return 'ods';
  if (lower.endsWith('.csv')) return 'csv';
  if (lower.endsWith('.tsv') || lower.endsWith('.tab')) return 'tsv';
  return 'xlsx';
}

/**
 * The full "open this file and make it the active workbook" flow used by both
 * File → Open and drag-and-drop. Parses the file, replaces the active
 * workbook, and replays imported hyperlinks once the new unit mounts (see
 * the AddHyperLinkCommand side-channel notes above).
 *
 * `replaceWorkbook` is the React-state setter from `WorkbookContext`. It has
 * to stay outside this module so the file is decoupled from React.
 */
export async function loadSpreadsheetFile(
  file: File,
  _api: FUniver | null,
  replaceWorkbook: (data: IWorkbookData, format: WorkbookFormat) => void,
  onProgress?: OpenProgress,
): Promise<void> {
  const data = await openSpreadsheetFile(file, onProgress);
  const format = inferFormat(file.name);
  onProgress?.('mounting');
  replaceWorkbook(data, format);
  // Hyperlinks are now baked into cell.p inline by the parser (see
  // parse-impl.ts `buildHyperlinkBody`). The previous side-channel
  // (`data.__pendingHyperlinks`) + per-link AddHyperLinkCommand replay
  // is gone — that was O(N) awaited round-trips per open.

  // Capture into the recent-files list so the landing screen can
  // surface this file next visit. Fire-and-forget — IDB failures are
  // non-fatal for the open path.
  void recordRecentFile({
    name: file.name,
    sourceFormat: format,
    data,
    size: file.size,
    openedAt: Date.now(),
  }).catch((err) => console.warn('[recent-files] capture failed', err));
}

export type SaveOptions = {
  /** Outline / group state to fold into the xlsx — see ExportExtras.outline. */
  outline?: OutlineState;
  /** Chart models to fold into the xlsx — see ExportExtras.charts. */
  charts?: ChartModel[];
  /** Pivot models to fold into the xlsx — see ExportExtras.pivots. */
  pivots?: PivotModel[];
};

export async function saveAsXlsx(
  api: FUniver,
  filename = 'workbook.xlsx',
  options: SaveOptions = {},
) {
  const wb = api.getActiveWorkbook();
  if (!wb) return;
  // wb.save() is a deep clone of the whole workbook — measurable on big
  // sheets. Capture it once and pass it down to anything that needs the
  // snapshot (xlsx writer + hyperlink extractor).
  const snapshot = timeIt('snapshot-save', () => wb.save() as IWorkbookData);
  const chartImages = await renderChartImagesForExport(api, options.charts);
  const extras: ExportExtras = {
    ...collectExportExtras(snapshot),
    ...(options.outline ? { outline: options.outline } : {}),
    ...(options.charts && options.charts.length > 0 ? { charts: options.charts } : {}),
    ...(chartImages.length > 0 ? { chartImages } : {}),
    ...(options.pivots && options.pivots.length > 0 ? { pivots: options.pivots } : {}),
  };
  const blob = await workbookDataToXlsx(snapshot, extras);
  const finalName = ensureExt(filename, 'xlsx');
  triggerDownload(blob, finalName);
  toast(api, `Saved as ${finalName}`);
  // The on-disk file now supersedes the autosave slot — drop it so a
  // crash after a successful save doesn't re-prompt with stale state.
  void discardAutosaveAfterExplicitSave();
}

/**
 * Snapshot the active workbook to xlsx bytes WITHOUT triggering a
 * download. Used by the co-edit share flow to upload the room's
 * starting workbook to the server. Returns null if no workbook is
 * mounted.
 */
export async function exportCurrentWorkbookAsXlsxBlob(
  api: FUniver,
  options: SaveOptions = {},
): Promise<Blob | null> {
  const wb = api.getActiveWorkbook();
  if (!wb) return null;
  const snapshot = timeIt('snapshot-save', () => wb.save() as IWorkbookData);
  const chartImages = await renderChartImagesForExport(api, options.charts);
  const extras: ExportExtras = {
    ...collectExportExtras(snapshot),
    ...(options.outline ? { outline: options.outline } : {}),
    ...(options.charts && options.charts.length > 0 ? { charts: options.charts } : {}),
    ...(chartImages.length > 0 ? { chartImages } : {}),
    ...(options.pivots && options.pivots.length > 0 ? { pivots: options.pivots } : {}),
  };
  return workbookDataToXlsx(snapshot, extras);
}

/**
 * Render every live chart to a PNG so the worker can embed it. Runs on
 * the main thread because ECharts needs a DOM container. Errors render
 * to null and are skipped — a missing image is better than a failed
 * export, and the JSON sidecar still carries the editable model.
 */
async function renderChartImagesForExport(
  api: FUniver,
  charts: ChartModel[] | undefined,
): Promise<NonNullable<ExportExtras['chartImages']>> {
  if (!charts?.length) return [];
  const out: NonNullable<ExportExtras['chartImages']> = [];
  for (const model of charts) {
    try {
      const { width, height } = pixelsForChart(api, model);
      const png = await renderChartToPng(api, model, width, height);
      if (!png) continue;
      out.push({ chartId: model.id, sheetId: model.sheetId, png, anchor: model.pos });
    } catch (err) {
      console.warn('[xlsx export] chart render failed; skipping', model.id, err);
    }
  }
  return out;
}

/**
 * Resolve Univer's IMessageService off the FUniver injector and show a brief
 * success toast. Silently no-ops if the message service isn't registered (in
 * tests, in headless seed paths, etc.) — feedback is nice-to-have, not load-
 * bearing.
 */
function toast(api: FUniver, content: string): void {
  // Dev-only record so e2e specs can verify the call without depending on
  // Sonner's lazily-mounted toast portal (which made the assertion flaky
  // on cold CI runners). Production builds tree-shake `import.meta.env.DEV`
  // away — this is a no-op cost in shipped bundles.
  if (import.meta.env.DEV) {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const sink = (globalThis as any).__toastLog__ as Array<{ content: string }> | undefined;
    if (Array.isArray(sink)) sink.push({ content });
    else
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (globalThis as any).__toastLog__ = [{ content }];
  }
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const injector = (api as any)._injector as
    | { get: (token: unknown) => unknown }
    | undefined;
  if (!injector) return;
  try {
    const svc = injector.get(IMessageService) as
      | { show: (opts: { content: string; type?: MessageType; duration?: number }) => unknown }
      | undefined;
    svc?.show({ content, type: MessageType.Success, duration: 2500 });
  } catch {
    /* message service not registered — silent */
  }
}

/**
 * Read hyperlinks out of the workbook snapshot. AddHyperLinkCommand stores
 * the URL in the cell's rich-text body (`cell.p.body.customRanges` with
 * `rangeType: HYPERLINK`), NOT in `HyperLinkModel` (the model is a sparse
 * index that the command notably does not populate). So the snapshot is the
 * source of truth — we just have to look inside `cell.p`, which the plain
 * xlsx exporter otherwise ignores.
 */
function collectExportExtras(snapshot: IWorkbookData): ExportExtras {
  return { hyperlinks: extractHyperlinks(snapshot) };
}

type HyperlinkExtra = { row: number; column: number; payload: string; display?: string };

function extractHyperlinks(
  snapshot: IWorkbookData,
): Record<string, HyperlinkExtra[]> {
  const out: Record<string, HyperlinkExtra[]> = {};
  for (const sheetId of snapshot.sheetOrder ?? []) {
    const wsd = snapshot.sheets?.[sheetId];
    if (!wsd?.cellData) continue;
    const links: HyperlinkExtra[] = [];
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
          const display = dataStream.slice(cr.startIndex, cr.endIndex + 1);
          links.push({ row: r, column: c, payload: url, display });
        }
      }
    }
    if (links.length) out[sheetId] = links;
  }
  return out;
}


export async function saveAsOds(api: FUniver, filename = 'workbook.ods') {
  const wb = api.getActiveWorkbook();
  if (!wb) return;
  const snapshot = timeIt('snapshot-save', () => wb.save() as IWorkbookData);
  const blob = await workbookDataToOds(snapshot);
  const finalName = ensureExt(filename, 'ods');
  triggerDownload(blob, finalName);
  toast(api, `Saved as ${finalName}`);
}

export async function saveAsCsv(api: FUniver, filename = 'workbook.csv') {
  const wb = api.getActiveWorkbook();
  if (!wb) return;
  const snapshot = timeIt('snapshot-save', () => wb.save() as IWorkbookData);
  const blob = await workbookDataToDelimited(snapshot, 'csv');
  const finalName = ensureExt(filename, 'csv');
  triggerDownload(blob, finalName);
  toast(api, `Saved as ${finalName}`);
}

export async function saveAsTsv(api: FUniver, filename = 'workbook.tsv') {
  const wb = api.getActiveWorkbook();
  if (!wb) return;
  const snapshot = timeIt('snapshot-save', () => wb.save() as IWorkbookData);
  const blob = await workbookDataToDelimited(snapshot, 'tsv');
  const finalName = ensureExt(filename, 'tsv');
  triggerDownload(blob, finalName);
  toast(api, `Saved as ${finalName}`);
}

function ensureExt(name: string, ext: string): string {
  const re = new RegExp(`\\.${ext}$`, 'i');
  return re.test(name) ? name : `${name.replace(/\.(xlsx|ods)$/i, '')}.${ext}`;
}

function triggerDownload(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  // Slight delay so the click handler completes before revoking.
  setTimeout(() => URL.revokeObjectURL(url), 0);
}

export function pickXlsxFile(): Promise<File | null> {
  console.info('[open-xlsx] opening file picker');
  return new Promise((resolve) => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept =
      '.xlsx,.ods,.csv,.tsv,.tab,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.oasis.opendocument.spreadsheet,text/csv,text/tab-separated-values';
    input.style.display = 'none';

    let settled = false;
    const settle = (file: File | null) => {
      if (settled) return;
      settled = true;
      input.remove();
      resolve(file);
    };

    input.addEventListener(
      'change',
      () => {
        const file = input.files?.[0] ?? null;
        console.info('[open-xlsx] file chosen', file?.name);
        settle(file);
      },
      { once: true },
    );
    // Standardized cancel event — fires when the user dismisses the native
    // dialog without picking a file. Replaces the older focus-based heuristic
    // which raced the change event in some browsers (resolving "cancelled"
    // before the file selection arrived).
    input.addEventListener(
      'cancel',
      () => {
        console.info('[open-xlsx] picker cancelled');
        settle(null);
      },
      { once: true },
    );

    document.body.appendChild(input);
    input.click();
  });
}
