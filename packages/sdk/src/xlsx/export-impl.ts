import ExcelJS from 'exceljs';
import { CustomRangeType, type IStyleData, type IWorkbookData } from '@univerjs/core';
import { univerStyleToExcel } from './style-mapping';
import { RESOURCES_SHEET } from './constants';
import { commentBodyToString, readCommentsFromSnapshot, refToRowCol } from './comments-resource';
import { applyPageSetupToXlsxWorksheet, readPageSetupFromSnapshot } from './page-setup-resource';
import {
  applyDataValidationToXlsxWorksheet,
  readDataValidationFromSnapshot,
} from './data-validation-resource';
import { applyTablesToXlsxWorksheet, readTablesFromSnapshot } from './tables-resource';
import {
  applyPassthroughToXlsxBuffer,
  mimeForPassthrough,
  readPassthroughFromSnapshot,
} from './passthrough-resource';
import type { ExportExtras } from './export';

/**
 * Pure conversion: Univer IWorkbookData → ExcelJS workbook → xlsx Blob.
 * Imported only by `exporter.worker.ts` (where it actually runs) — the
 * library `export.ts` is the worker-dispatch entry point. Splitting like
 * this keeps ExcelJS (~600 KB) out of the main chunk.
 *
 * Feature models that don't live on `IWorkbookData` (charts / pivots /
 * sparklines / outline resources) are baked into the snapshot by the host
 * BEFORE calling the exporter; this core only reads what's on the snapshot
 * plus the generic `ExportExtras` (hyperlinks, outline gutter, images).
 */

type ICellSnapshot = {
  v?: string | number | boolean;
  f?: string;
  s?: string | IStyleData;
};

const PX_PER_CHAR = 7;
const pxToChars = (px: number) => Math.max(0, px / PX_PER_CHAR);
const pxToPoints = (px: number) => Math.max(0, (px * 72) / 96);

export async function workbookDataToXlsxImpl(
  data: IWorkbookData,
  extras: ExportExtras = {},
): Promise<Blob> {
  const wb = new ExcelJS.Workbook();
  wb.title = data.name || 'Untitled';

  // Map our `custom.properties` slot (set by File → Properties) into
  // ExcelJS workbook properties so the file's Core + App properties
  // (docProps/core.xml + docProps/app.xml) carry them. Without this
  // mapping the Properties dialog values would silently drop on every
  // save round-trip; we'd also leave `created` / `modified` blank,
  // which makes the file look untouched to Windows Explorer / Finder.
  const customProps = (data.custom?.properties ?? {}) as {
    title?: string;
    subject?: string;
    author?: string;
    tags?: string;
    category?: string;
    description?: string;
    company?: string;
    manager?: string;
    createdAt?: string;
    modifiedAt?: string;
  };
  if (customProps.title) wb.title = customProps.title;
  if (customProps.subject) wb.subject = customProps.subject;
  if (customProps.description) wb.description = customProps.description;
  // Excel's "Tags" / "Keywords" field is a single comma-separated string.
  if (customProps.tags) wb.keywords = customProps.tags;
  if (customProps.category) wb.category = customProps.category;
  if (customProps.company) wb.company = customProps.company;
  if (customProps.manager) wb.manager = customProps.manager;
  if (customProps.author) {
    wb.creator = customProps.author;
    wb.lastModifiedBy = customProps.author;
  }
  // Created — fall back to now if absent so an unset file at least has
  // a non-1900 timestamp. Modified always bumps to the moment of save.
  const createdIso = customProps.createdAt ?? new Date().toISOString();
  wb.created = new Date(createdIso);
  wb.modified = new Date();

  const resolveStyle = (s: ICellSnapshot['s']): IStyleData | undefined => {
    if (!s) return undefined;
    if (typeof s === 'string') return (data.styles?.[s] ?? undefined) as IStyleData | undefined;
    return s as IStyleData;
  };

  // Read the thread-comment resource once — the per-sheet loop just
  // indexes into the resulting map. Cheap when no comments exist
  // (the helper returns `{}` without touching JSON.parse).
  const commentsBySheet = readCommentsFromSnapshot(data);
  // Same pre-parse for page setup so each sheet just looks up its
  // entry below — keeps the JSON.parse out of the per-sheet loop.
  const pageSetupBySheet = readPageSetupFromSnapshot(data);
  // Data-validation rules live in Univer's plugin resource. We pull
  // them up-front and apply per-sheet below so a file opened in real
  // Excel keeps its list / whole / date / etc. constraints — and the
  // round-trip via our pipeline doesn't quietly drop them.
  const dataValidationBySheet = readDataValidationFromSnapshot(data);
  // Tables (xlsx ListObjects) round-trip through a passthrough
  // sidecar — Univer doesn't model them as first-class objects, so we
  // re-add via ExcelJS's `addTable` on save.
  const tablesBySheet = readTablesFromSnapshot(data);

  for (const sheetId of data.sheetOrder) {
    const wsd = data.sheets[sheetId];
    if (!wsd) continue;
    const ws = wb.addWorksheet(wsd.name ?? sheetId);

    if (wsd.hidden === 1) ws.state = 'hidden';

    if (wsd.tabColor && typeof wsd.tabColor === 'string' && wsd.tabColor.startsWith('#')) {
      const rgb = wsd.tabColor.slice(1).toUpperCase();
      if (/^[0-9A-F]{6}$/.test(rgb)) {
        ws.properties.tabColor = { argb: `FF${rgb}` };
      }
    }

    if (wsd.freeze && (wsd.freeze.xSplit > 0 || wsd.freeze.ySplit > 0)) {
      ws.views = [
        {
          state: 'frozen',
          xSplit: wsd.freeze.xSplit || 0,
          ySplit: wsd.freeze.ySplit || 0,
        },
      ];
    }

    if (typeof wsd.defaultColumnWidth === 'number' && wsd.defaultColumnWidth > 0) {
      ws.properties.defaultColWidth = pxToChars(wsd.defaultColumnWidth);
    }
    if (typeof wsd.defaultRowHeight === 'number' && wsd.defaultRowHeight > 0) {
      ws.properties.defaultRowHeight = pxToPoints(wsd.defaultRowHeight);
    }

    const cellData = (wsd.cellData ?? {}) as Record<string, Record<string, ICellSnapshot>>;
    for (const rKey of Object.keys(cellData)) {
      const r = Number(rKey);
      const row = cellData[rKey];
      for (const cKey of Object.keys(row)) {
        const c = Number(cKey);
        const cell = row[cKey];

        const excelCell = ws.getCell(r + 1, c + 1);

        if (cell.f) {
          const formula = cell.f.startsWith('=') ? cell.f.slice(1) : cell.f;
          excelCell.value = { formula, result: cell.v ?? null } as ExcelJS.CellValue;
        } else if (cell.v !== undefined && cell.v !== null) {
          excelCell.value = cell.v as ExcelJS.CellValue;
        }

        // Hyperlink encoded by the parser into `cell.p.body.customRanges`
        // (the shape sheets-hyper-link's AddHyperLinkCommand writes).
        // Promote that to the ExcelJS-native `{ text, hyperlink }`
        // value so a foreign reader (real Excel, gsheets) sees a live
        // link instead of just the display text. extras.hyperlinks
        // below still wins for the live save path — this branch only
        // matters when the exporter is called without extras (audit
        // round-trip, headless seed-back).
        const cellP = (cell as ICellSnapshot & { p?: unknown }).p as
          | {
              body?: {
                dataStream?: string;
                customRanges?: Array<{ rangeType?: number; properties?: { url?: string } }>;
              };
            }
          | undefined;
        if (cellP?.body?.customRanges) {
          for (const cr of cellP.body.customRanges) {
            if (
              cr.rangeType === CustomRangeType.HYPERLINK &&
              typeof cr.properties?.url === 'string' &&
              cr.properties.url
            ) {
              const display =
                cellP.body.dataStream?.replace(/[\r\n]+$/, '') ??
                (typeof cell.v === 'string' ? cell.v : String(cell.v ?? ''));
              excelCell.value = {
                text: display,
                hyperlink: cr.properties.url,
              } as ExcelJS.CellValue;
              break;
            }
          }
        }

        const styleObj = resolveStyle(cell.s);
        if (styleObj) {
          Object.assign(excelCell, univerStyleToExcel(styleObj));
        }
      }
    }

    if (Array.isArray(wsd.mergeData)) {
      for (const m of wsd.mergeData) {
        ws.mergeCells(m.startRow + 1, m.startColumn + 1, m.endRow + 1, m.endColumn + 1);
      }
    }

    // Page-setup passthrough — orientation / paper / header-footer
    // text the parser stashed. Apply BEFORE comments so anything in
    // the header that references page state lands in a consistent
    // worksheet object (ExcelJS is OK either way; ordering is for
    // readers of this file).
    const ps = pageSetupBySheet[sheetId];
    if (ps) applyPageSetupToXlsxWorksheet(ws, ps);

    // Thread-comment resource → xlsx-native cell notes. Comments
    // round-trip through `data.resources` for our own re-open
    // (Univer's plugin loads them) AND via `excelCell.note` so a
    // file opened in real Excel renders the same yellow triangle +
    // pop-up. The JSON parse happens once below the loop start;
    // each sheet just looks up its bucket.
    const sheetComments = commentsBySheet[sheetId] ?? [];
    for (const c of sheetComments) {
      const { row, column } = refToRowCol(c.ref);
      if (row < 0 || column < 0) continue;
      const excelCell = ws.getCell(row + 1, column + 1);
      const text = commentBodyToString(c.text);
      if (!text) continue;
      // ExcelJS accepts a string for the simple note case. The
      // expanded `{ texts: [{ text }] }` form lets us carry rich
      // styling later; not used today because the audit only
      // verifies text and the import path doesn't preserve runs.
      (excelCell as unknown as { note: string }).note = text;
    }

    // Data validation rules — re-apply so list dropdowns / whole-number
    // / etc. constraints survive Save→Open in Excel. Skipped for sheets
    // without any rules to avoid touching ExcelJS's lazy model.
    const dvRules = dataValidationBySheet[sheetId];
    if (dvRules?.length) applyDataValidationToXlsxWorksheet(ws, dvRules);

    // Tables (xlsx ListObjects). Re-added before extras.hyperlinks so
    // any link written into a table cell still overrides the table's
    // initial value below.
    const tablesForSheet = tablesBySheet[sheetId];
    if (tablesForSheet?.length) applyTablesToXlsxWorksheet(ws, tablesForSheet);

    const sheetHyperlinks = extras.hyperlinks?.[sheetId] ?? [];
    for (const hl of sheetHyperlinks) {
      const excelCell = ws.getCell(hl.row + 1, hl.column + 1);
      const text =
        hl.display ??
        (typeof excelCell.value === 'string'
          ? excelCell.value
          : typeof excelCell.value === 'number'
            ? String(excelCell.value)
            : hl.payload);
      excelCell.value = { text, hyperlink: hl.payload };
    }

    const columnData = (wsd.columnData ?? {}) as Record<string, { w?: number; hd?: number }>;
    for (const cKey of Object.keys(columnData)) {
      const c = Number(cKey);
      const meta = columnData[cKey];
      if (typeof meta?.w === 'number' && meta.w > 0) {
        ws.getColumn(c + 1).width = pxToChars(meta.w);
      }
      if (meta?.hd === 1) ws.getColumn(c + 1).hidden = true;
    }

    const rowData = (wsd.rowData ?? {}) as Record<string, { h?: number; hd?: number }>;
    for (const rKey of Object.keys(rowData)) {
      const r = Number(rKey);
      const meta = rowData[rKey];
      if (typeof meta?.h === 'number' && meta.h > 0) {
        ws.getRow(r + 1).height = pxToPoints(meta.h);
      }
      if (meta?.hd === 1) ws.getRow(r + 1).hidden = true;
    }

    const sheetOutline = extras.outline?.[sheetId];
    if (sheetOutline) {
      for (const g of sheetOutline.rows ?? []) {
        for (let r = g.start; r <= g.end; r++) {
          const row = ws.getRow(r + 1);
          row.outlineLevel = 1;
          if (g.collapsed) row.hidden = true;
        }
      }
      for (const g of sheetOutline.cols ?? []) {
        for (let c = g.start; c <= g.end; c++) {
          const col = ws.getColumn(c + 1);
          col.outlineLevel = 1;
          if (g.collapsed) col.hidden = true;
        }
      }
    }

    // Pre-rendered bitmaps embedded as floating images (e.g. chart P5b chart
    // snapshots). The live model still ships in the snapshot resources so our
    // app re-attaches an editable object on re-open; a foreign reader sees the
    // image. `editAs: 'oneCell'` anchors to the top-left cell only.
    const sheetCharts = (extras.chartImages ?? []).filter((c) => c.sheetId === sheetId);
    for (const ci of sheetCharts) {
      const imageId = wb.addImage({ buffer: ci.png, extension: 'png' });
      // ExcelJS's runtime accepts simple `{col, row}` anchors (per its
      // docs), but the published type insists on the full `Anchor`
      // struct. Cast to bypass — runtime is the source of truth.
      ws.addImage(imageId, {
        tl: { col: ci.anchor.startColumn, row: ci.anchor.startRow },
        br: { col: ci.anchor.endColumn + 1, row: ci.anchor.endRow + 1 },
        editAs: 'oneCell',
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
      } as any);
    }
  }

  // Lift Univer's defined-name resource into xlsx-native `<definedName>`
  // entries so the named ranges survive a round-trip through Microsoft
  // Excel (which ignores our hidden sidecar). We still ship the sidecar
  // below for our own re-open path; this is the foreign-reader leg.
  writeDefinedNamesToXlsx(wb, data);

  if (Array.isArray(data.resources) && data.resources.length > 0) {
    const meta = wb.addWorksheet(RESOURCES_SHEET);
    meta.state = 'veryHidden';
    const json = JSON.stringify(data.resources);
    const CHUNK = 30_000;
    for (let i = 0, row = 1; i < json.length; i += CHUNK, row++) {
      meta.getCell(row, 1).value = json.slice(i, i + CHUNK);
    }
  }

  const buf = await wb.xlsx.writeBuffer();

  // Re-inject raw OOXML parts captured at parse time (today: VBA macros).
  // No-op when the snapshot has no passthrough sidecar; ExcelJS-written
  // buffer is returned as-is in that path.
  const passthrough = readPassthroughFromSnapshot(data);
  const patched = await applyPassthroughToXlsxBuffer(buf as ArrayBuffer, passthrough);

  return new Blob([patched], { type: mimeForPassthrough(passthrough) });
}

const DEFINED_NAMES_RESOURCE = 'SHEET_DEFINED_NAME_PLUGIN';
function writeDefinedNamesToXlsx(wb: ExcelJS.Workbook, data: IWorkbookData): void {
  const res = data.resources?.find((r) => r.name === DEFINED_NAMES_RESOURCE);
  if (!res?.data) return;
  let map: Record<string, { name?: string; formulaOrRefString?: string }>;
  try {
    map = JSON.parse(res.data);
  } catch {
    return;
  }
  for (const entry of Object.values(map ?? {})) {
    const name = entry?.name;
    const ref = entry?.formulaOrRefString;
    if (!name || !ref) continue;
    // Multi-range defined names come back as a comma-joined string —
    // split and add each separately so ExcelJS encodes them correctly.
    for (const piece of ref.split(',')) {
      const cleaned = piece.trim();
      if (!cleaned) continue;
      try {
        wb.definedNames.add(cleaned, name);
      } catch {
        /* Invalid ref strings (unbounded, references to gone sheets,
         * etc.) shouldn't kill the export — skip the bad entry. */
      }
    }
  }
}
