import type { IStyleData } from '@univerjs/core';
import type * as ExcelJS from 'exceljs';

/**
 * Mappings between ExcelJS and Univer style models.
 *
 * Univer's `IStyleData` uses two-letter compact keys (bl, it, ff, fs, cl, bg,
 * ht, vt, n, ul, bd) plus numeric enums for alignment. ExcelJS uses verbose
 * keys (font.bold, alignment.horizontal, numFmt, etc.). Round-trip preserves
 * the basics; advanced styling (gradients, patterns, complex borders) falls
 * back to defaults.
 */

const ARGB_RX = /^#?([0-9A-Fa-f]{6,8})$/;

function normalizeColor(argb: string | undefined): string | undefined {
  if (!argb) return undefined;
  const m = ARGB_RX.exec(argb);
  if (!m) return undefined;
  const hex = m[1];
  // ExcelJS uses ARGB; drop alpha if present.
  const rgb = hex.length === 8 ? hex.slice(2) : hex;
  return `#${rgb.toLowerCase()}`;
}

// Accept Univer's Nullable<string> (string | null | void) by treating any
// non-string input as undefined.
function toARGB(rgb: unknown): string | undefined {
  if (typeof rgb !== 'string' || !rgb) return undefined;
  const m = ARGB_RX.exec(rgb);
  if (!m) return undefined;
  const hex = m[1];
  const norm = hex.length === 8 ? hex : `FF${hex}`;
  return norm.toUpperCase();
}

// Univer HorizontalAlign: 0 = default, 1 = LEFT, 2 = CENTER, 3 = RIGHT
const H_ALIGN_FROM_EXCEL: Record<string, number> = {
  left: 1,
  center: 2,
  right: 3,
};
const H_ALIGN_TO_EXCEL: Record<number, string> = {
  1: 'left',
  2: 'center',
  3: 'right',
};

// Univer VerticalAlign: 0 = default, 1 = TOP, 2 = MIDDLE, 3 = BOTTOM
const V_ALIGN_FROM_EXCEL: Record<string, number> = {
  top: 1,
  middle: 2,
  bottom: 3,
};
const V_ALIGN_TO_EXCEL: Record<number, string> = {
  1: 'top',
  2: 'middle',
  3: 'bottom',
};

export function excelStyleToUniver(cell: ExcelJS.Cell): IStyleData | undefined {
  const s: IStyleData = {};

  if (cell.font) {
    if (cell.font.name) s.ff = cell.font.name;
    if (cell.font.size) s.fs = cell.font.size;
    if (cell.font.bold) s.bl = 1;
    if (cell.font.italic) s.it = 1;
    if (cell.font.underline) s.ul = { s: 1 };
    const fc = normalizeColor((cell.font.color as { argb?: string } | undefined)?.argb);
    if (fc) s.cl = { rgb: fc };
  }

  if (cell.fill && cell.fill.type === 'pattern' && cell.fill.pattern === 'solid') {
    const bg = normalizeColor((cell.fill.fgColor as { argb?: string } | undefined)?.argb);
    if (bg) s.bg = { rgb: bg };
  }

  if (cell.alignment) {
    const ha = cell.alignment.horizontal;
    if (ha && H_ALIGN_FROM_EXCEL[ha] !== undefined) s.ht = H_ALIGN_FROM_EXCEL[ha];
    const va = cell.alignment.vertical;
    if (va && V_ALIGN_FROM_EXCEL[va] !== undefined) s.vt = V_ALIGN_FROM_EXCEL[va];
    if (cell.alignment.wrapText) s.tb = 3; // WrapStrategy.WRAP
  }

  if (cell.numFmt) s.n = { pattern: cell.numFmt };

  // Borders — keep simple: any side present => thin border, color from spec.
  if (cell.border) {
    const bd: NonNullable<IStyleData['bd']> = {};
    const sides: Array<['t' | 'b' | 'l' | 'r', 'top' | 'bottom' | 'left' | 'right']> = [
      ['t', 'top'],
      ['b', 'bottom'],
      ['l', 'left'],
      ['r', 'right'],
    ];
    for (const [k, key] of sides) {
      const side = (cell.border as Record<string, unknown>)[key] as
        | { style?: string; color?: { argb?: string } }
        | undefined;
      if (side?.style && side.style !== 'none') {
        bd[k] = { s: 1, cl: { rgb: normalizeColor(side.color?.argb) ?? '#666666' } };
      }
    }
    if (Object.keys(bd).length > 0) s.bd = bd;
  }

  return Object.keys(s).length > 0 ? s : undefined;
}

export function univerStyleToExcel(style: IStyleData): Partial<ExcelJS.Style> {
  const out: Partial<ExcelJS.Style> = {};

  const font: ExcelJS.Style['font'] = {};
  if (style.ff) font.name = style.ff;
  if (style.fs) font.size = style.fs;
  if (style.bl === 1) font.bold = true;
  if (style.it === 1) font.italic = true;
  if (style.ul?.s === 1) font.underline = true;
  if (style.cl && typeof style.cl === 'object' && 'rgb' in style.cl) {
    const argb = toARGB(style.cl.rgb);
    if (argb) font.color = { argb };
  }
  if (Object.keys(font).length > 0) out.font = font;

  if (style.bg && typeof style.bg === 'object' && 'rgb' in style.bg) {
    const argb = toARGB(style.bg.rgb);
    if (argb) {
      out.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb },
      };
    }
  }

  const alignment: ExcelJS.Style['alignment'] = {};
  if (style.ht && H_ALIGN_TO_EXCEL[style.ht]) {
    alignment.horizontal = H_ALIGN_TO_EXCEL[style.ht] as ExcelJS.Style['alignment']['horizontal'];
  }
  if (style.vt && V_ALIGN_TO_EXCEL[style.vt]) {
    alignment.vertical = V_ALIGN_TO_EXCEL[style.vt] as ExcelJS.Style['alignment']['vertical'];
  }
  if (style.tb === 3) alignment.wrapText = true;
  if (Object.keys(alignment).length > 0) out.alignment = alignment;

  if (style.n?.pattern) out.numFmt = style.n.pattern;

  if (style.bd) {
    const border: ExcelJS.Style['border'] = {};
    const sides: Array<['t' | 'b' | 'l' | 'r', 'top' | 'bottom' | 'left' | 'right']> = [
      ['t', 'top'],
      ['b', 'bottom'],
      ['l', 'left'],
      ['r', 'right'],
    ];
    for (const [k, key] of sides) {
      const side = style.bd[k];
      if (!side) continue;
      const color = side.cl && 'rgb' in side.cl ? toARGB(side.cl.rgb) : undefined;
      (border as Record<string, unknown>)[key] = {
        style: 'thin',
        ...(color ? { color: { argb: color } } : {}),
      };
    }
    if (Object.keys(border).length > 0) out.border = border;
  }

  return out;
}
