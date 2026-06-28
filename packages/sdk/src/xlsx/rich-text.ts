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

/**
 * In-cell rich text ↔ xlsx round-trip (T4.2).
 *
 * Excel stores a cell with mixed per-character formatting as a "rich text"
 * value: an array of runs, each with its own font. ExcelJS surfaces that as
 * `cell.value = { richText: [{ text, font }, …] }`. Univer represents the same
 * thing as `cell.p` — an `IDocumentBody` whose `textRuns[]` carry per-range
 * styles over a `dataStream`.
 *
 * The SDK previously dropped this both ways (import flattened `richText` to a
 * plain string; export had no `textRuns` path), so opening an Excel file with a
 * bold word in a cell and saving it back lost the formatting. These pure
 * mappers bridge the two. Kept @univerjs-value-free (types only, erased at
 * runtime) so they're unit-testable under the node:test + tsx runner.
 */

/** One ExcelJS rich-text run. */
export type ExcelRichRun = {
  text: string;
  font?: {
    name?: string;
    size?: number;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    strike?: boolean;
    color?: { argb?: string };
  };
};

/** Minimal `IDocumentBody` shape we read/write for cell rich text. */
export type RichBody = {
  dataStream: string;
  textRuns?: Array<{ st: number; ed: number; ts?: Record<string, unknown> }>;
  paragraphs?: Array<{ startIndex: number }>;
  sectionBreaks?: Array<{ startIndex: number }>;
};

const ARGB_RX = /^[0-9a-fA-F]{6,8}$/;

function argbToRgb(argb: string | undefined): string | undefined {
  if (!argb) return undefined;
  if (!ARGB_RX.test(argb)) return undefined;
  const hex = argb.length === 8 ? argb.slice(2) : argb;
  return `#${hex.toLowerCase()}`;
}

function rgbToArgb(rgb: unknown): string | undefined {
  if (typeof rgb !== 'string') return undefined;
  const hex = rgb.startsWith('#') ? rgb.slice(1) : rgb;
  if (!ARGB_RX.test(hex)) return undefined;
  return (hex.length === 6 ? `FF${hex}` : hex).toUpperCase();
}

/** ExcelJS run font → Univer text style (compact keys). Empty object if plain. */
function fontToTextStyle(font: ExcelRichRun['font']): Record<string, unknown> {
  const ts: Record<string, unknown> = {};
  if (!font) return ts;
  if (font.name) ts.ff = font.name;
  if (font.size) ts.fs = font.size;
  if (font.bold) ts.bl = 1;
  if (font.italic) ts.it = 1;
  if (font.underline) ts.ul = { s: 1 };
  if (font.strike) ts.st = { s: 1 };
  const rgb = argbToRgb(font.color?.argb);
  if (rgb) ts.cl = { rgb };
  return ts;
}

/** Univer text style → ExcelJS run font. Undefined if no styling. */
function textStyleToFont(
  ts: Record<string, unknown> | undefined,
): ExcelRichRun['font'] | undefined {
  if (!ts) return undefined;
  const font: NonNullable<ExcelRichRun['font']> = {};
  if (typeof ts.ff === 'string') font.name = ts.ff;
  if (typeof ts.fs === 'number') font.size = ts.fs;
  if (ts.bl === 1) font.bold = true;
  if (ts.it === 1) font.italic = true;
  if ((ts.ul as { s?: number } | undefined)?.s === 1) font.underline = true;
  if ((ts.st as { s?: number } | undefined)?.s === 1) font.strike = true;
  const argb = rgbToArgb((ts.cl as { rgb?: unknown } | undefined)?.rgb);
  if (argb) font.color = { argb };
  return Object.keys(font).length > 0 ? font : undefined;
}

/** True if at least one run carries real per-run formatting worth preserving. */
function hasFormatting(runs: ExcelRichRun[]): boolean {
  return runs.some((r) => Object.keys(fontToTextStyle(r.font)).length > 0);
}

/**
 * ExcelJS `richText` runs → an `IDocumentBody`. Returns `undefined` when there's
 * nothing worth preserving (no run carries formatting) so the caller keeps the
 * cheap plain-string path. Mirrors the fork's `transformTextNodes2Document`
 * dataStream convention (text + `\r\n`, one paragraph, one section break).
 */
export function excelRichTextToBody(runs: ExcelRichRun[]): RichBody | undefined {
  if (!Array.isArray(runs) || runs.length === 0) return undefined;
  if (!hasFormatting(runs)) return undefined;
  let stream = '';
  const textRuns: NonNullable<RichBody['textRuns']> = [];
  for (const run of runs) {
    const text = run.text ?? '';
    if (text.length === 0) continue;
    const st = stream.length;
    stream += text;
    const ts = fontToTextStyle(run.font);
    textRuns.push(
      Object.keys(ts).length > 0 ? { st, ed: stream.length, ts } : { st, ed: stream.length },
    );
  }
  if (textRuns.length === 0) return undefined;
  stream += '\r\n';
  return {
    dataStream: stream,
    textRuns,
    paragraphs: [{ startIndex: stream.length - 2 }],
    sectionBreaks: [{ startIndex: stream.length - 1 }],
  };
}

/**
 * An `IDocumentBody` → ExcelJS `richText` runs. Returns `undefined` unless the
 * body actually has >1 styled run (a single uniform run is better expressed as
 * a plain value + cell-level style, the existing path). The trailing `\r\n` is
 * stripped from the emitted text.
 */
export function bodyToExcelRichText(body: RichBody | undefined): ExcelRichRun[] | undefined {
  const runs = body?.textRuns;
  if (!body || !runs || runs.length === 0) return undefined;
  const stream = body.dataStream ?? '';
  const end = stream.replace(/[\r\n]+$/, '').length;
  const out: ExcelRichRun[] = [];
  let cursor = 0;
  const sorted = [...runs].sort((a, b) => a.st - b.st);
  for (const r of sorted) {
    const st = Math.max(0, Math.min(r.st, end));
    const ed = Math.max(st, Math.min(r.ed, end));
    // Gap before this run → an unstyled run so no text is lost.
    if (st > cursor) out.push({ text: stream.slice(cursor, st) });
    const text = stream.slice(st, ed);
    if (text.length > 0) {
      const font = textStyleToFont(r.ts);
      out.push(font ? { text, font } : { text });
    }
    cursor = ed;
  }
  if (cursor < end) out.push({ text: stream.slice(cursor, end) });
  const meaningful = out.filter((r) => r.text.length > 0);
  // Only worth a richText value if there's real per-run formatting.
  if (meaningful.length === 0 || !meaningful.some((r) => r.font)) return undefined;
  return meaningful;
}
