import type { FUniver } from '@univerjs/core/facade';

/** Direction for the selection to move after commit. Mirrors Excel's
 *  Enter / Shift+Enter / Tab / Shift+Tab behavior in the formula bar. */
export type CommitDirection = 'down' | 'up' | 'right' | 'left' | 'none';

/**
 * Commit a string the user typed into the formula bar to the active cell.
 *   - Strings starting with `=` are treated as formulas.
 *   - Numeric strings are coerced to numbers so the cell isn't stored as text.
 *   - Empty string clears the cell.
 *
 * After write, optionally move the active cell by one in `direction`.
 * Clamped at the sheet edges (no wrap; Excel doesn't wrap either).
 */
export function commitToActiveCell(
  api: FUniver,
  raw: string,
  direction: CommitDirection = 'none',
) {
  const sheet = api.getActiveWorkbook()?.getActiveSheet();
  const range = sheet?.getActiveRange();
  if (!range) return;
  const row = range.getRow();
  const col = range.getColumn();
  const cell = sheet!.getRange(row, col);

  const text = raw.trim();
  if (text === '') {
    cell.setValue({ v: null });
  } else if (text.startsWith('=')) {
    cell.setValue({ f: text });
  } else if (!Number.isNaN(Number(text))) {
    // Coerce pure numbers — Univer treats string-typed numbers as text otherwise.
    cell.setValue({ v: Number(text) });
  } else {
    cell.setValue({ v: text });
  }

  if (direction === 'none') return;
  const nextRow = direction === 'down' ? row + 1 : direction === 'up' ? Math.max(0, row - 1) : row;
  const nextCol =
    direction === 'right' ? col + 1 : direction === 'left' ? Math.max(0, col - 1) : col;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (sheet as any).getRange(nextRow, nextCol)?.activate?.();
}

/**
 * Cycle the absolute/relative prefix on the cell reference under the
 * caret — Excel's F4 behavior. Cycle order: A1 → $A$1 → A$1 → $A1 → A1.
 *
 * Returns the rewritten value and the new caret position, or null if
 * no reference was found under the caret.
 *
 * Sheet-qualified refs (`Sheet2!A1`, `'My Sheet'!A1`) are detected and
 * only the cell-ref portion after the `!` is rewritten — the sheet
 * qualifier is preserved verbatim. A range (`A1:B2`) cycles only the
 * endpoint under the caret, matching Excel.
 */
export function cycleAbsoluteRefAtCaret(
  value: string,
  caret: number,
): { value: string; caret: number } | null {
  // A1-style ref with optional $ on either component, optional sheet
  // qualifier. We scan for any ref overlapping the caret position.
  const refRe = /(?:'[^']+'|[A-Za-z_][\w.]*)?!?(\$?)([A-Za-z]{1,3})(\$?)(\d+)/g;
  let m: RegExpExecArray | null;
  while ((m = refRe.exec(value)) !== null) {
    const start = m.index;
    const end = start + m[0].length;
    // Caret can be inside OR immediately after the ref — matches Excel.
    if (caret < start || caret > end) continue;
    const [whole, colAbs, colLetters, rowAbs, rowDigits] = m;
    // Strip the sheet qualifier (incl. trailing !) from the rewrite —
    // we only cycle the cell-ref part.
    const bang = whole.lastIndexOf('!');
    const prefix = bang >= 0 ? whole.slice(0, bang + 1) : '';
    const next = nextAbsCombo(colAbs === '$', rowAbs === '$');
    const rewritten = `${prefix}${next.col ? '$' : ''}${colLetters}${next.row ? '$' : ''}${rowDigits}`;
    const out = value.slice(0, start) + rewritten + value.slice(end);
    return { value: out, caret: start + rewritten.length };
  }
  return null;
}

function nextAbsCombo(colAbs: boolean, rowAbs: boolean): { col: boolean; row: boolean } {
  // A1 → $A$1 → A$1 → $A1 → A1
  if (!colAbs && !rowAbs) return { col: true, row: true };
  if (colAbs && rowAbs) return { col: false, row: true };
  if (!colAbs && rowAbs) return { col: true, row: false };
  return { col: false, row: false };
}

/**
 * Quote a sheet name for use in a formula reference. Excel requires
 * single quotes around any name that's not a plain identifier (spaces,
 * dashes, leading digits, etc.). Embedded apostrophes are doubled.
 */
export function quoteSheetName(name: string): string {
  if (/^[A-Za-z_][A-Za-z0-9_]*$/.test(name)) return name;
  return `'${name.replace(/'/g, "''")}'`;
}

/**
 * Insert a cell/range reference at the caret in a formula draft,
 * replacing any trailing ref token if one is there. This is how
 * Excel's range-picker mode behaves: each canvas click while a
 * formula is being typed replaces the most recently-typed ref.
 *
 * Examples:
 *   ("=SUM(", 5, "A1")        → "=SUM(A1"
 *   ("=SUM(A1", 7, "B2")      → "=SUM(B2"      (replaces trailing A1)
 *   ("=SUM(A1+", 8, "B2")     → "=SUM(A1+B2"   (operator before, append)
 *   ("=SUM(Sheet2!A1", N, ..) → replaces the whole qualified ref token
 */
export function insertRefAtCaret(
  value: string,
  caret: number,
  ref: string,
): { value: string; caret: number } {
  const before = value.slice(0, caret);
  const after = value.slice(caret);
  // Match a trailing ref token at the end of `before` — optional
  // sheet qualifier (quoted or bare) + `!`, then `$?col$?row` with an
  // optional `:end` half.
  const trailRe = /(?:'[^']*'|[A-Za-z_][A-Za-z0-9_]*)?!?\$?[A-Za-z]+\$?\d+(?::\$?[A-Za-z]+\$?\d+)?$/;
  const m = before.match(trailRe);
  const start = m ? caret - m[0].length : caret;
  const next = value.slice(0, start) + ref + after;
  return { value: next, caret: start + ref.length };
}

/**
 * True when the caret in a formula draft is at a position where typing
 * a cell reference makes sense — start of the formula, or right after
 * an operator / open paren / comma. Used to gate the range-picker so
 * clicks during e.g. `=SUM(A1, "literal` don't hijack the literal.
 */
export function canInsertRefAtCaret(value: string, caret: number): boolean {
  if (!value.startsWith('=')) return false;
  // Disallow when caret sits inside a quoted string literal — count
  // unescaped double quotes before caret; odd = inside a string.
  let inString = false;
  for (let i = 1; i < caret && i < value.length; i++) {
    if (value[i] === '"') inString = !inString;
  }
  if (inString) return false;
  return true;
}
