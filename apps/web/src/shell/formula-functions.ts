/**
 * Canonical list of formula functions exposed in autocomplete. Mirrors the
 * Formulas-tab categories so power users can also discover via the ribbon.
 *
 * Each entry has a short description used by the formula-bar autocomplete
 * popover. Keep names UPPER-CASE.
 */
export type FormulaFn = {
  name: string;
  description: string;
  args?: string[];
};

export const FORMULA_FUNCTIONS: FormulaFn[] = [
  // Math & Trig
  { name: 'SUM', description: 'Adds its arguments', args: ['number1', '[number2]'] },
  { name: 'SUMIF', description: 'Adds cells specified by a given criteria', args: ['range', 'criteria', '[sum_range]'] },
  { name: 'SUMIFS', description: 'Adds cells matching multiple criteria', args: ['sum_range', 'criteria_range1', 'criteria1', '[criteria_range2]', '[criteria2]'] },
  { name: 'ABS', description: 'Returns the absolute value', args: ['number'] },
  { name: 'INT', description: 'Rounds down to the nearest integer', args: ['number'] },
  { name: 'MOD', description: 'Returns the remainder', args: ['number', 'divisor'] },
  { name: 'POWER', description: 'Returns a number raised to a power', args: ['number', 'power'] },
  { name: 'ROUND', description: 'Rounds a number to a specified number of digits', args: ['number', 'num_digits'] },
  { name: 'ROUNDUP', description: 'Rounds a number up', args: ['number', 'num_digits'] },
  { name: 'ROUNDDOWN', description: 'Rounds a number down', args: ['number', 'num_digits'] },
  { name: 'SQRT', description: 'Returns the square root', args: ['number'] },
  { name: 'PRODUCT', description: 'Multiplies its arguments', args: ['number1', '[number2]'] },

  // Statistical
  { name: 'AVERAGE', description: 'Returns the average of its arguments', args: ['number1', '[number2]'] },
  { name: 'AVERAGEIF', description: 'Average of cells matching a criterion', args: ['range', 'criteria', '[average_range]'] },
  { name: 'COUNT', description: 'Counts the number of numbers', args: ['value1', '[value2]'] },
  { name: 'COUNTA', description: 'Counts non-empty cells', args: ['value1', '[value2]'] },
  { name: 'COUNTIF', description: 'Counts cells matching a criterion', args: ['range', 'criteria'] },
  { name: 'COUNTIFS', description: 'Counts cells matching multiple criteria', args: ['criteria_range1', 'criteria1', '[criteria_range2]', '[criteria2]'] },
  { name: 'MAX', description: 'Returns the maximum value', args: ['number1', '[number2]'] },
  { name: 'MIN', description: 'Returns the minimum value', args: ['number1', '[number2]'] },
  { name: 'MEDIAN', description: 'Returns the median', args: ['number1', '[number2]'] },
  { name: 'STDEV', description: 'Estimates standard deviation', args: ['number1', '[number2]'] },

  // Text
  { name: 'CONCATENATE', description: 'Joins several text items into one', args: ['text1', '[text2]'] },
  { name: 'CONCAT', description: 'Joins several text items into one', args: ['text1', '[text2]'] },
  { name: 'LEFT', description: 'Returns the leftmost characters', args: ['text', '[num_chars]'] },
  { name: 'RIGHT', description: 'Returns the rightmost characters', args: ['text', '[num_chars]'] },
  { name: 'MID', description: 'Returns characters from the middle', args: ['text', 'start_num', 'num_chars'] },
  { name: 'LEN', description: 'Returns the length of a text string', args: ['text'] },
  { name: 'UPPER', description: 'Converts text to uppercase', args: ['text'] },
  { name: 'LOWER', description: 'Converts text to lowercase', args: ['text'] },
  { name: 'PROPER', description: 'Capitalises each word', args: ['text'] },
  { name: 'TRIM', description: 'Removes excess whitespace', args: ['text'] },
  { name: 'SUBSTITUTE', description: 'Replaces occurrences of a substring', args: ['text', 'old_text', 'new_text', '[instance_num]'] },
  { name: 'REPLACE', description: 'Replaces a range of characters', args: ['old_text', 'start_num', 'num_chars', 'new_text'] },
  { name: 'FIND', description: 'Finds text within text (case-sensitive)', args: ['find_text', 'within_text', '[start_num]'] },
  { name: 'SEARCH', description: 'Finds text within text (case-insensitive)', args: ['find_text', 'within_text', '[start_num]'] },

  // Date & Time
  { name: 'TODAY', description: "Returns today's date", args: [] },
  { name: 'NOW', description: 'Returns the current date and time', args: [] },
  { name: 'DATE', description: 'Returns a date from y/m/d arguments', args: ['year', 'month', 'day'] },
  { name: 'YEAR', description: 'Returns the year', args: ['serial_number'] },
  { name: 'MONTH', description: 'Returns the month', args: ['serial_number'] },
  { name: 'DAY', description: 'Returns the day of the month', args: ['serial_number'] },
  { name: 'WEEKDAY', description: 'Returns the day of the week', args: ['serial_number', '[return_type]'] },
  { name: 'DATEDIF', description: 'Calculates the difference between two dates', args: ['start_date', 'end_date', 'unit'] },

  // Logical
  { name: 'IF', description: 'Specifies a logical test', args: ['logical_test', 'value_if_true', 'value_if_false'] },
  { name: 'IFS', description: 'Multiple conditions', args: ['logical_test1', 'value_if_true1', '[logical_test2]', '[value_if_true2]'] },
  { name: 'AND', description: 'Returns TRUE only if all arguments are TRUE', args: ['logical1', '[logical2]'] },
  { name: 'OR', description: 'Returns TRUE if any argument is TRUE', args: ['logical1', '[logical2]'] },
  { name: 'NOT', description: 'Reverses the logical value', args: ['logical'] },
  { name: 'IFERROR', description: 'Returns a value if expression is an error', args: ['value', 'value_if_error'] },

  // Lookup & Reference
  { name: 'VLOOKUP', description: 'Vertical lookup', args: ['lookup_value', 'table_array', 'col_index_num', '[range_lookup]'] },
  { name: 'HLOOKUP', description: 'Horizontal lookup', args: ['lookup_value', 'table_array', 'row_index_num', '[range_lookup]'] },
  { name: 'INDEX', description: 'Returns a value from a position', args: ['array', 'row_num', '[column_num]'] },
  { name: 'MATCH', description: 'Returns the position of a value', args: ['lookup_value', 'lookup_array', '[match_type]'] },
  { name: 'INDIRECT', description: 'Returns the reference specified by a string', args: ['ref_text', '[a1]'] },
  { name: 'CHOOSE', description: 'Chooses a value from a list', args: ['index_num', 'value1', '[value2]'] },
  { name: 'OFFSET', description: 'Returns a reference offset from a given reference', args: ['reference', 'rows', 'cols', '[height]', '[width]'] },
];

const FN_BY_NAME = new Map(FORMULA_FUNCTIONS.map((fn) => [fn.name, fn] as const));

const FN_BY_PREFIX_INDEX = (() => {
  const m = new Map<string, FormulaFn[]>();
  for (const fn of FORMULA_FUNCTIONS) {
    for (let i = 1; i <= fn.name.length; i++) {
      const key = fn.name.slice(0, i);
      const arr = m.get(key) ?? [];
      arr.push(fn);
      m.set(key, arr);
    }
  }
  return m;
})();

/**
 * Return the function suggestions matching `prefix` (case-insensitive),
 * sorted alphabetically. Empty prefix → empty result.
 */
export function suggestFunctions(prefix: string, limit = 8): FormulaFn[] {
  if (!prefix) return [];
  const upper = prefix.toUpperCase();
  return (FN_BY_PREFIX_INDEX.get(upper) ?? []).slice(0, limit);
}

export function getFunctionByName(name: string): FormulaFn | null {
  return FN_BY_NAME.get(name.toUpperCase()) ?? null;
}

export function getFunctionSignature(fn: FormulaFn): string {
  return fn.args && fn.args.length > 0 ? `${fn.name}(${fn.args.join(', ')})` : `${fn.name}()`;
}

/**
 * Match active workbook sheet names against a prefix. Excel-style
 * autocomplete: typing `=Sa` in the formula bar shows `Sales` (and
 * any other sheets starting with "Sa") alongside the function
 * suggestions. Returns the names case-insensitively filtered, in
 * the order they appear on the workbook.
 *
 * Sheets containing characters outside `[A-Za-z0-9_]` need to be
 * quoted in formulas (e.g. `'My Sheet'!A1`); v0.1.1 only suggests
 * names that the fragment extractor can match (alphanumeric +
 * underscore). Names with spaces / symbols are still type-able
 * manually with the apostrophe quoting Excel uses, just not
 * auto-completed yet.
 */
export function suggestSheetNames(
  prefix: string,
  sheetNames: readonly string[],
  limit = 5,
): string[] {
  if (!prefix) return [];
  const lower = prefix.toLowerCase();
  const out: string[] = [];
  for (const name of sheetNames) {
    if (name.toLowerCase().startsWith(lower) && /^[A-Za-z0-9_]+$/.test(name)) {
      out.push(name);
      if (out.length >= limit) break;
    }
  }
  return out;
}

/**
 * Inspect the formula-bar value at caret position and return the
 * "in-flight" function name fragment, if any. E.g.:
 *   "=SU" at caret 3 → "SU"
 *   "=SUM(A1)+AVE" at caret 13 → "AVE"
 *   "=1+2" → null (no in-flight name)
 *   "Hello" → null (no formula)
 */
export function extractFunctionFragment(text: string, caret: number): string | null {
  if (!text.startsWith('=')) return null;
  // Walk back from caret over [A-Z0-9_] characters.
  let i = caret;
  while (i > 0 && /[A-Za-z0-9_]/.test(text[i - 1] ?? '')) i--;
  const frag = text.slice(i, caret);
  if (!frag || /^\d/.test(frag)) return null;
  return frag;
}

export function findFunctionContextAtCaret(
  text: string,
  caret: number,
): { fn: FormulaFn; openParen: number; closeParen: number | null } | null {
  if (!text.startsWith('=')) return null;
  let depth = 0;
  for (let i = Math.min(caret - 1, text.length - 1); i >= 0; i--) {
    const ch = text[i];
    if (ch === ')') {
      depth += 1;
      continue;
    }
    if (ch !== '(') continue;
    if (depth > 0) {
      depth -= 1;
      continue;
    }
    let j = i;
    while (j > 0 && /[A-Za-z0-9_]/.test(text[j - 1] ?? '')) j--;
    const name = text.slice(j, i);
    const fn = getFunctionByName(name);
    if (!fn) return null;
    let closeParen: number | null = null;
    let innerDepth = 0;
    for (let k = i + 1; k < text.length; k++) {
      if (text[k] === '(') innerDepth += 1;
      else if (text[k] === ')') {
        if (innerDepth === 0) {
          closeParen = k;
          break;
        }
        innerDepth -= 1;
      }
    }
    return { fn, openParen: i, closeParen };
  }
  return null;
}

export function expandFunctionArgsAtCaret(
  text: string,
  caret: number,
): { value: string; selectionStart: number; selectionEnd: number } | null {
  const ctx = findFunctionContextAtCaret(text, caret);
  if (!ctx) return null;
  const args = ctx.fn.args ?? [];
  if (args.length === 0) return null;
  const template = args.join(', ');
  const innerStart = ctx.openParen + 1;
  const innerEnd = ctx.closeParen ?? caret;
  const existing = text.slice(innerStart, innerEnd).trim();
  if (existing.length > 0) return null;
  const suffix = ctx.closeParen === null ? ')' : '';
  const value = `${text.slice(0, innerStart)}${template}${suffix}${text.slice(innerEnd)}`;
  return {
    value,
    selectionStart: innerStart,
    selectionEnd: innerStart + template.length,
  };
}
