import ExcelJS from 'exceljs';
import type { IStyleData, IWorkbookData } from '@univerjs/core';
import { univerStyleToExcel } from './style-mapping';
import { writeOutlineIntoSnapshot } from '../outline/resources';
import { writeChartsIntoSnapshot } from '../charts/resources';
import { writePivotsIntoSnapshot } from '../pivots/resources';
import { RESOURCES_SHEET } from './constants';
import type { ExportExtras } from './export';

/**
 * Pure conversion: Univer IWorkbookData → ExcelJS workbook → xlsx Blob.
 * Imported only by `exporter.worker.ts` (where it actually runs) — the
 * main bundle's `export.ts` is type-only + the worker-dispatch entry
 * point. Splitting like this keeps ExcelJS (~600 KB) out of the main
 * chunk.
 *
 * Behavior is byte-for-byte equivalent to the previous in-place impl;
 * see export.ts header for the fidelity scope.
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
      ws.views = [{
        state: 'frozen',
        xSplit: wsd.freeze.xSplit || 0,
        ySplit: wsd.freeze.ySplit || 0,
      }];
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

    // Charts P5b — embed pre-rendered chart bitmaps as floating images.
    // The live chart model still ships via `__casual_sheets_charts__` so
    // our app re-attaches an editable chart on re-open; Excel sees the
    // image. `editAs: 'oneCell'` anchors the image so it scales with the
    // top-left cell only (Excel's "Move with cells but don't size").
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

  if (extras.outline && Object.keys(extras.outline).length > 0) {
    writeOutlineIntoSnapshot(data, extras.outline);
  }

  if (extras.charts && extras.charts.length > 0) {
    writeChartsIntoSnapshot(data, extras.charts);
  }

  if (extras.pivots && extras.pivots.length > 0) {
    writePivotsIntoSnapshot(data, extras.pivots);
  }

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
  return new Blob([buf], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
}
