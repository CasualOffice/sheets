import type { IWorkbookData } from '@univerjs/core';
import type { FUniver } from '@univerjs/core/facade';
import { workbookDataToXlsx, xlsxToWorkbookData } from '../xlsx';

/**
 * File-level imperative actions. Pure functions — the caller owns React state
 * (e.g. lifting the workbook snapshot so a new Open replaces the active unit).
 */

export async function openXlsx(file: File): Promise<IWorkbookData> {
  console.info('[open-xlsx] reading file', { name: file.name, size: file.size });
  const buf = await file.arrayBuffer();
  console.info('[open-xlsx] buffer read', buf.byteLength, 'bytes — parsing');
  const data = await xlsxToWorkbookData(buf);
  console.info('[open-xlsx] parsed', { id: data.id, sheets: Object.keys(data.sheets ?? {}).length });
  data.name = file.name.replace(/\.xlsx$/i, '');
  return data;
}

export async function saveAsXlsx(api: FUniver, filename = 'workbook.xlsx') {
  const wb = api.getActiveWorkbook();
  if (!wb) return;
  const snapshot = wb.save() as IWorkbookData;
  const blob = await workbookDataToXlsx(snapshot);
  triggerDownload(blob, ensureXlsxExt(filename));
}

function ensureXlsxExt(name: string): string {
  return /\.xlsx$/i.test(name) ? name : `${name}.xlsx`;
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
    input.accept = '.xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
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
