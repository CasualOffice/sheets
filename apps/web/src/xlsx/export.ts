import ExcelJS from 'exceljs';
import type { IStyleData, IWorkbookData } from '@univerjs/core';
import { univerStyleToExcel } from './style-mapping';

/**
 * Convert a Univer `IWorkbookData` snapshot to an .xlsx Blob.
 * See `import.ts` for the fidelity scope (same coverage in both directions).
 */
export async function workbookDataToXlsx(data: IWorkbookData): Promise<Blob> {
  const wb = new ExcelJS.Workbook();
  wb.title = data.name || 'Untitled';

  // Resolve a style ref (string id or inline IStyleData) to the IStyleData object.
  const resolveStyle = (s: ICellSnapshot['s']): IStyleData | undefined => {
    if (!s) return undefined;
    if (typeof s === 'string') return (data.styles?.[s] ?? undefined) as IStyleData | undefined;
    return s as IStyleData;
  };

  for (const sheetId of data.sheetOrder) {
    const wsd = data.sheets[sheetId];
    if (!wsd) continue;
    const ws = wb.addWorksheet(wsd.name ?? sheetId);

    const cellData = (wsd.cellData ?? {}) as Record<string, Record<string, ICellSnapshot>>;
    for (const rKey of Object.keys(cellData)) {
      const r = Number(rKey);
      const row = cellData[rKey];
      for (const cKey of Object.keys(row)) {
        const c = Number(cKey);
        const cell = row[cKey];

        // ExcelJS uses 1-indexed positions.
        const excelCell = ws.getCell(r + 1, c + 1);

        if (cell.f) {
          // Formula cell — strip leading '=' (ExcelJS adds it back).
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

    // Merges
    if (Array.isArray(wsd.mergeData)) {
      for (const m of wsd.mergeData) {
        ws.mergeCells(m.startRow + 1, m.startColumn + 1, m.endRow + 1, m.endColumn + 1);
      }
    }
  }

  const buf = await wb.xlsx.writeBuffer();
  return new Blob([buf], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
}

type ICellSnapshot = {
  v?: string | number | boolean;
  f?: string;
  s?: string | IStyleData;
};
