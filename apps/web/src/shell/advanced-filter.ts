/**
 * Advanced Filter — pure criteria matching (Excel's Data → Advanced).
 *
 * Excel's criteria-range model: the criteria range has a header row of field
 * names matching the list's headers; each subsequent row is a set of conditions
 * ANDed across its columns, and the rows are ORed together. A list row passes if
 * it satisfies any one criteria row. Criterion cells use Excel's text/number
 * grammar: a comparison operator (`>`, `<`, `>=`, `<=`, `<>`, `=`) followed by a
 * value, or a bare value (numbers → equality; text → "begins with", the Excel
 * default), with `*` / `?` wildcards (escape with `~`).
 *
 * Pure + framework-free so it's unit-testable; the dialog reads the three ranges
 * with getValues() and writes the matching rows to the "copy to" location.
 */

export type FilterValue = string | number | boolean | null | undefined;

const toText = (v: FilterValue): string => (v == null ? '' : String(v));
const asNumber = (v: FilterValue): number | null => {
  if (typeof v === 'number') return v;
  if (typeof v === 'string' && v.trim() !== '' && Number.isFinite(Number(v))) return Number(v);
  return null;
};

/** Compile an Excel wildcard pattern (`*`, `?`, `~`-escaped) to a RegExp. */
function wildcardToRegExp(pattern: string, anchorEnd: boolean): RegExp {
  let out = '';
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i];
    if (ch === '~' && i + 1 < pattern.length) {
      out += pattern[i + 1].replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      i++;
    } else if (ch === '*') {
      out += '.*';
    } else if (ch === '?') {
      out += '.';
    } else {
      out += ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }
  }
  return new RegExp(`^${out}${anchorEnd ? '$' : ''}`, 'i');
}

/** Build a predicate for one criterion string against a cell value. */
export function compileCriterion(raw: string): (cell: FilterValue) => boolean {
  const crit = raw.trim();
  if (crit === '') return () => true; // empty cell matches everything

  const opMatch = /^(>=|<=|<>|>|<|=)/.exec(crit);
  const op = opMatch?.[1];
  const rest = op ? crit.slice(op.length).trim() : crit;
  const restNum = asNumber(rest);

  if (op === '>' || op === '<' || op === '>=' || op === '<=') {
    return (cell) => {
      const n = asNumber(cell ?? null);
      if (n == null || restNum == null) {
        // Fall back to string comparison when either side isn't numeric.
        const a = toText(cell);
        if (op === '>') return a > rest;
        if (op === '<') return a < rest;
        if (op === '>=') return a >= rest;
        return a <= rest;
      }
      if (op === '>') return n > restNum;
      if (op === '<') return n < restNum;
      if (op === '>=') return n >= restNum;
      return n <= restNum;
    };
  }

  if (op === '<>') {
    const re = wildcardToRegExp(rest, true);
    return (cell) => !re.test(toText(cell));
  }

  if (op === '=') {
    // Exact match (wildcards still honoured).
    const re = wildcardToRegExp(rest, true);
    return (cell) => re.test(toText(cell));
  }

  // Bare criterion: numbers → equality; text → begins-with (Excel's default).
  if (restNum != null) {
    return (cell) => asNumber(cell ?? null) === restNum;
  }
  const re = wildcardToRegExp(rest, false); // not end-anchored → "begins with"
  return (cell) => re.test(toText(cell));
}

export interface AdvancedFilterInput {
  listHeader: FilterValue[];
  listRows: FilterValue[][];
  critHeader: FilterValue[];
  critRows: FilterValue[][];
}

/**
 * Return the indices (into `listRows`) of rows matching the criteria. A row
 * matches if it satisfies any criteria row (OR); within a criteria row all
 * non-empty cells must match their column (AND). Criteria columns are matched to
 * list columns by header name (case-insensitive); unknown headers are ignored.
 */
export function matchAdvancedFilter(input: AdvancedFilterInput): number[] {
  const { listHeader, listRows, critHeader, critRows } = input;
  const colOf = new Map<string, number>();
  listHeader.forEach((h, i) => colOf.set(toText(h).trim().toLowerCase(), i));

  // Pre-compile each criteria row into [listColIndex, predicate] pairs.
  const compiledRows = critRows.map((cr) => {
    const conds: Array<{ col: number; test: (c: FilterValue) => boolean }> = [];
    cr.forEach((cell, ci) => {
      const text = toText(cell).trim();
      if (text === '') return; // blank criterion = no constraint
      const field = toText(critHeader[ci]).trim().toLowerCase();
      const col = colOf.get(field);
      if (col == null) return; // criterion column not in the list → ignore
      conds.push({ col, test: compileCriterion(text) });
    });
    return conds;
  });

  // A criteria row with no usable conditions matches everything (Excel).
  const matches = (row: FilterValue[]): boolean =>
    compiledRows.some((conds) => conds.every(({ col, test }) => test(row[col])));

  const out: number[] = [];
  // No criteria rows at all → nothing to filter on → match nothing.
  if (compiledRows.length === 0) return out;
  listRows.forEach((row, i) => {
    if (matches(row)) out.push(i);
  });
  return out;
}
