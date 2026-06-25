/**
 * PDF export — "Download as PDF" (File menu). Renders the active sheet's used
 * range as a real vector table via jsPDF + jspdf-autotable: searchable text,
 * auto-paginated across A4 pages, with column letters + row numbers and a
 * title. Values come from the workbook snapshot (same `.v` the print path uses);
 * rich per-cell styling (fonts/fills/borders) is a follow-up — this ships the
 * MUST-have one-click PDF that OnlyOffice/Google Sheets have and we lacked.
 */
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import type { FUniver } from '@univerjs/core/facade';
import { saveWorkbook } from '../univer-facade';

export type PdfResult = { ok: true } | { ok: false; reason: 'empty' | 'cancelled' };

type DeskBridge = {
  isDesktop?: boolean;
  saveAs(suggestedName: string, bytes: ArrayBuffer): Promise<string | null>;
};

function deskBridge(): DeskBridge | undefined {
  if (typeof window === 'undefined') return undefined;
  const b = (window as unknown as { __deskApp__?: DeskBridge }).__deskApp__;
  return b?.isDesktop ? b : undefined;
}

/** 0-based column index → spreadsheet letter (0→A, 25→Z, 26→AA). */
function colLetter(n: number): string {
  let s = '';
  let x = n + 1;
  while (x > 0) {
    const m = (x - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    x = Math.floor((x - 1) / 26);
  }
  return s;
}

function cellText(v: unknown): string {
  if (v === undefined || v === null) return '';
  return typeof v === 'string' ? v : String(v);
}

export async function exportActiveSheetPdf(api: FUniver): Promise<PdfResult> {
  const wb = api.getActiveWorkbook();
  const ws = wb?.getActiveSheet();
  if (!wb || !ws) return { ok: false, reason: 'empty' };

  const snap = saveWorkbook(wb);
  const sId = (ws as { getSheetId?: () => string }).getSheetId?.();
  const sheet = sId ? snap?.sheets?.[sId] : undefined;
  const cellData = (sheet?.cellData ?? {}) as Record<string, Record<string, { v?: unknown }>>;

  // Used range from populated cells.
  let maxRow = -1;
  let maxCol = -1;
  for (const rk of Object.keys(cellData)) {
    const row = cellData[rk];
    for (const ck of Object.keys(row)) {
      const v = row[ck]?.v;
      if (v === undefined || v === null || v === '') continue;
      const r = Number(rk);
      const c = Number(ck);
      if (r > maxRow) maxRow = r;
      if (c > maxCol) maxCol = c;
    }
  }
  if (maxRow < 0 || maxCol < 0) return { ok: false, reason: 'empty' };

  // Resolve per-cell style (bold / horizontal align / fill) so the PDF reflects
  // the sheet's look. `cell.s` is a style id into the workbook style table, or
  // an inline style object. Univer HorizontalAlign: 1=left, 2=center, 3=right.
  type RawStyle = { bl?: number; ht?: number; bg?: { rgb?: string } };
  const wbStyles = (snap?.styles ?? {}) as Record<string, RawStyle>;
  const HALIGN: Record<number, 'left' | 'center' | 'right'> = { 1: 'left', 2: 'center', 3: 'right' };
  type CellStyle = { bold: boolean; halign?: 'left' | 'center' | 'right'; fill?: string };

  const head = [['', ...Array.from({ length: maxCol + 1 }, (_, c) => colLetter(c))]];
  const body: string[][] = [];
  const styleGrid: CellStyle[][] = [];
  for (let r = 0; r <= maxRow; r++) {
    const cells: string[] = [String(r + 1)];
    const styles: CellStyle[] = [];
    const row = cellData[r] as Record<string, { v?: unknown; s?: unknown }> | undefined;
    for (let c = 0; c <= maxCol; c++) {
      cells.push(cellText(row?.[c]?.v));
      const sRef = row?.[c]?.s;
      const st = (typeof sRef === 'string' ? wbStyles[sRef] : (sRef as RawStyle | undefined)) ?? undefined;
      styles.push({
        bold: st?.bl === 1,
        halign: typeof st?.ht === 'number' ? HALIGN[st.ht] : undefined,
        fill: st?.bg?.rgb,
      });
    }
    body.push(cells);
    styleGrid.push(styles);
  }

  const wbName = (wb as { getName?: () => string }).getName?.() ?? 'Workbook';
  const sName = (ws as { getSheetName?: () => string }).getSheetName?.() ?? 'Sheet';

  // Wide sheets read better in landscape.
  const doc = new jsPDF({ orientation: maxCol > 6 ? 'landscape' : 'portrait', unit: 'pt', format: 'a4' });
  doc.setFontSize(14);
  doc.text(`${wbName} — ${sName}`, 40, 36);
  doc.setFontSize(9);
  doc.setTextColor(120);
  doc.text(new Date().toLocaleString(), 40, 50);
  doc.setTextColor(0);

  autoTable(doc, {
    head,
    body,
    startY: 62,
    styles: { fontSize: 8, cellPadding: 3, overflow: 'linebreak', lineColor: [203, 213, 225], lineWidth: 0.5 },
    headStyles: { fillColor: [241, 245, 249], textColor: 30, fontStyle: 'bold' },
    // Row-number gutter column.
    columnStyles: { 0: { fillColor: [241, 245, 249], textColor: 90, halign: 'right', cellWidth: 30 } },
    theme: 'grid',
    margin: { top: 62, left: 40, right: 40, bottom: 40 },
    // Apply per-cell bold / alignment / fill from the sheet.
    didParseCell: (data) => {
      if (data.section !== 'body') return;
      const col = data.column.index;
      if (col === 0) return; // row-number gutter
      const sg = styleGrid[data.row.index]?.[col - 1];
      if (!sg) return;
      if (sg.bold) data.cell.styles.fontStyle = 'bold';
      if (sg.halign) data.cell.styles.halign = sg.halign;
      if (sg.fill) data.cell.styles.fillColor = sg.fill;
    },
  });

  const file = `${wbName}-${sName}`.replace(/[^\w.-]+/g, '_');
  const filename = `${file}.pdf`;
  // Desktop shell: Export must open the native picker (and never a phantom
  // browser download — the shell's hard rule). Route the PDF bytes through
  // the bridge's Save As so the user sees and chooses where it lands. The web
  // build keeps jsPDF's own download.
  const bridge = deskBridge();
  if (bridge) {
    const bytes = doc.output('arraybuffer') as ArrayBuffer;
    const saved = await bridge.saveAs(filename, bytes);
    return saved == null ? { ok: false, reason: 'cancelled' } : { ok: true };
  }
  doc.save(filename);
  return { ok: true };
}
