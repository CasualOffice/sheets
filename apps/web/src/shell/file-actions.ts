import type { IWorkbookData } from '@univerjs/core';
import type { FUniver } from '@univerjs/core/facade';
import { workbookDataToXlsx, xlsxToWorkbookData } from '../xlsx';

/**
 * File-level imperative actions. Pure functions — the caller owns React state
 * (e.g. lifting the workbook snapshot so a new Open replaces the active unit).
 */

export async function openXlsx(file: File): Promise<IWorkbookData> {
  const buf = await file.arrayBuffer();
  const data = await xlsxToWorkbookData(buf);
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
  return new Promise((resolve) => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    input.style.display = 'none';

    let resolved = false;
    const cleanup = () => {
      input.remove();
      window.removeEventListener('focus', onFocus);
    };
    const onChange = () => {
      resolved = true;
      const file = input.files?.[0] ?? null;
      cleanup();
      resolve(file);
    };
    // If the user cancels the OS file dialog there's no `change` event —
    // detect that by listening for focus returning to the window.
    const onFocus = () => {
      setTimeout(() => {
        if (resolved) return;
        cleanup();
        resolve(null);
      }, 200);
    };

    input.addEventListener('change', onChange, { once: true });
    window.addEventListener('focus', onFocus);
    document.body.appendChild(input);
    input.click();
  });
}
