import type { FUniver } from '@univerjs/core/facade';

/**
 * Print the active sheet via the browser's native print dialog.
 *
 * Univer OSS doesn't ship a print renderer (page setup / print preview are
 * Pro-only). We do the cheap-but-correct thing: render the active sheet's
 * used range as a plain HTML table styled for print, inject it into a hidden
 * iframe, and trigger window.print() in that iframe. The browser handles
 * paper size, margins, page headers/footers natively.
 *
 * Loss: formulas show their evaluated value (Excel's normal print behavior),
 * column widths are approximated, drawings/charts are not included.
 */
export function printActiveSheet(api: FUniver): void {
  const wb = api.getActiveWorkbook();
  if (!wb) return;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const fws: any = wb.getActiveSheet();
  if (!fws) return;
  const ws = fws.getSheet();

  // Find the used range so we don't print 1024 empty rows.
  const lastRow = ws.getLastRowWithContent?.() ?? 0;
  const lastCol = ws.getLastColumnWithContent?.() ?? 0;
  if (lastRow < 0 || lastCol < 0) return;

  const sheetName: string = fws.getSheetName?.() ?? 'Sheet';
  const workbookName: string = (wb as { getName?: () => string }).getName?.() ?? 'Workbook';

  const rows: string[] = [];
  for (let r = 0; r <= lastRow; r++) {
    const cells: string[] = [];
    for (let c = 0; c <= lastCol; c++) {
      const cell = ws.getCell(r, c);
      const raw: unknown = cell?.v;
      const text =
        raw === undefined || raw === null
          ? ''
          : typeof raw === 'string'
            ? raw
            : String(raw);
      cells.push(`<td>${escapeHtml(text)}</td>`);
    }
    rows.push(`<tr>${cells.join('')}</tr>`);
  }

  // Render column-width hints in CSS so the print layout doesn't squash
  // narrow columns. ExcelJS-style units don't matter here — relative
  // proportions are what carry through to print.
  const colWidths: string[] = [];
  for (let c = 0; c <= lastCol; c++) {
    const w = ws.getColumnWidth?.(c) ?? 88;
    colWidths.push(`<col style="width:${Math.round(w)}px" />`);
  }

  const html = `<!doctype html>
<html>
<head>
<meta charset="utf-8" />
<title>${escapeHtml(workbookName)} — ${escapeHtml(sheetName)}</title>
<style>
  @page { margin: 12mm; }
  body { font: 11pt Inter, system-ui, sans-serif; color: #1a1a1a; margin: 0; }
  h1 { font-size: 14pt; margin: 0 0 8pt; font-weight: 600; }
  .meta { color: #6b7280; font-size: 9pt; margin-bottom: 8pt; }
  table { border-collapse: collapse; width: 100%; }
  td { border: 1px solid #cbd5e1; padding: 4px 6px; vertical-align: top;
       font-size: 10pt; word-wrap: break-word; }
  thead td { background: #f1f5f9; font-weight: 600; }
  tr:nth-child(even) td { background: #f9fafb; }
</style>
</head>
<body>
  <h1>${escapeHtml(workbookName)} — ${escapeHtml(sheetName)}</h1>
  <div class="meta">Printed ${new Date().toLocaleString()}</div>
  <table>
    <colgroup>${colWidths.join('')}</colgroup>
    <tbody>${rows.join('')}</tbody>
  </table>
</body>
</html>`;

  // Hidden iframe avoids leaving the user's tab. Some browsers block
  // printing inside an iframe whose contentDocument has cross-origin
  // restrictions; writing via `srcdoc` keeps everything same-origin.
  const frame = document.createElement('iframe');
  frame.style.position = 'fixed';
  frame.style.right = '0';
  frame.style.bottom = '0';
  frame.style.width = '0';
  frame.style.height = '0';
  frame.style.border = '0';
  frame.srcdoc = html;
  document.body.appendChild(frame);

  frame.addEventListener(
    'load',
    () => {
      try {
        const win = frame.contentWindow;
        if (!win) return;
        win.focus();
        win.print();
      } finally {
        // Give the browser a moment to open its print dialog before we
        // detach the iframe; some engines abort the dialog otherwise.
        setTimeout(() => frame.remove(), 1000);
      }
    },
    { once: true },
  );
}

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
