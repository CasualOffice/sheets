import type { IWorkbookData } from '@univerjs/core';
import type { FUniver } from '@univerjs/core/facade';
import { workbookDataToXlsx, xlsxToWorkbookData } from '../xlsx';
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

/**
 * Open a spreadsheet from disk. We auto-detect by file extension and
 * dispatch to the right parser. xlsx files go through ExcelJS; ods files
 * through SheetJS Community.
 */
export async function openSpreadsheetFile(file: File): Promise<IWorkbookData> {
  console.info('[open] reading file', { name: file.name, size: file.size });
  const buf = await file.arrayBuffer();
  console.info('[open] buffer read', buf.byteLength, 'bytes — parsing');
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

export async function saveAsXlsx(api: FUniver, filename = 'workbook.xlsx') {
  const wb = api.getActiveWorkbook();
  if (!wb) return;
  const snapshot = wb.save() as IWorkbookData;
  const blob = await workbookDataToXlsx(snapshot);
  triggerDownload(blob, ensureExt(filename, 'xlsx'));
}

export async function saveAsOds(api: FUniver, filename = 'workbook.ods') {
  const wb = api.getActiveWorkbook();
  if (!wb) return;
  const snapshot = wb.save() as IWorkbookData;
  const blob = await workbookDataToOds(snapshot);
  triggerDownload(blob, ensureExt(filename, 'ods'));
}

export async function saveAsCsv(api: FUniver, filename = 'workbook.csv') {
  const wb = api.getActiveWorkbook();
  if (!wb) return;
  const snapshot = wb.save() as IWorkbookData;
  const blob = await workbookDataToDelimited(snapshot, 'csv');
  triggerDownload(blob, ensureExt(filename, 'csv'));
}

export async function saveAsTsv(api: FUniver, filename = 'workbook.tsv') {
  const wb = api.getActiveWorkbook();
  if (!wb) return;
  const snapshot = wb.save() as IWorkbookData;
  const blob = await workbookDataToDelimited(snapshot, 'tsv');
  triggerDownload(blob, ensureExt(filename, 'tsv'));
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
