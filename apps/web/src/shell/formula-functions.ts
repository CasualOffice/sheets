/**
 * Canonical list of formula functions exposed in autocomplete. Mirrors the
 * Formulas-tab categories so power users can also discover via the ribbon.
 *
 * Each entry has a short description used by the formula-bar autocomplete
 * popover. Keep names UPPER-CASE.
 */
export type FormulaFn = { name: string; description: string };

export const FORMULA_FUNCTIONS: FormulaFn[] = [
  // Math & Trig
  { name: 'SUM', description: 'Adds its arguments' },
  { name: 'SUMIF', description: 'Adds cells specified by a given criteria' },
  { name: 'SUMIFS', description: 'Adds cells matching multiple criteria' },
  { name: 'ABS', description: 'Returns the absolute value' },
  { name: 'INT', description: 'Rounds down to the nearest integer' },
  { name: 'MOD', description: 'Returns the remainder' },
  { name: 'POWER', description: 'Returns a number raised to a power' },
  { name: 'ROUND', description: 'Rounds a number to a specified number of digits' },
  { name: 'ROUNDUP', description: 'Rounds a number up' },
  { name: 'ROUNDDOWN', description: 'Rounds a number down' },
  { name: 'SQRT', description: 'Returns the square root' },
  { name: 'PRODUCT', description: 'Multiplies its arguments' },

  // Statistical
  { name: 'AVERAGE', description: 'Returns the average of its arguments' },
  { name: 'AVERAGEIF', description: 'Average of cells matching a criterion' },
  { name: 'COUNT', description: 'Counts the number of numbers' },
  { name: 'COUNTA', description: 'Counts non-empty cells' },
  { name: 'COUNTIF', description: 'Counts cells matching a criterion' },
  { name: 'COUNTIFS', description: 'Counts cells matching multiple criteria' },
  { name: 'MAX', description: 'Returns the maximum value' },
  { name: 'MIN', description: 'Returns the minimum value' },
  { name: 'MEDIAN', description: 'Returns the median' },
  { name: 'STDEV', description: 'Estimates standard deviation' },

  // Text
  { name: 'CONCATENATE', description: 'Joins several text items into one' },
  { name: 'CONCAT', description: 'Joins several text items into one' },
  { name: 'LEFT', description: 'Returns the leftmost characters' },
  { name: 'RIGHT', description: 'Returns the rightmost characters' },
  { name: 'MID', description: 'Returns characters from the middle' },
  { name: 'LEN', description: 'Returns the length of a text string' },
  { name: 'UPPER', description: 'Converts text to uppercase' },
  { name: 'LOWER', description: 'Converts text to lowercase' },
  { name: 'PROPER', description: 'Capitalises each word' },
  { name: 'TRIM', description: 'Removes excess whitespace' },
  { name: 'SUBSTITUTE', description: 'Replaces occurrences of a substring' },
  { name: 'REPLACE', description: 'Replaces a range of characters' },
  { name: 'FIND', description: 'Finds text within text (case-sensitive)' },
  { name: 'SEARCH', description: 'Finds text within text (case-insensitive)' },

  // Date & Time
  { name: 'TODAY', description: "Returns today's date" },
  { name: 'NOW', description: 'Returns the current date and time' },
  { name: 'DATE', description: 'Returns a date from y/m/d arguments' },
  { name: 'YEAR', description: 'Returns the year' },
  { name: 'MONTH', description: 'Returns the month' },
  { name: 'DAY', description: 'Returns the day of the month' },
  { name: 'WEEKDAY', description: 'Returns the day of the week' },
  { name: 'DATEDIF', description: 'Calculates the difference between two dates' },

  // Logical
  { name: 'IF', description: 'Specifies a logical test' },
  { name: 'IFS', description: 'Multiple conditions' },
  { name: 'AND', description: 'Returns TRUE only if all arguments are TRUE' },
  { name: 'OR', description: 'Returns TRUE if any argument is TRUE' },
  { name: 'NOT', description: 'Reverses the logical value' },
  { name: 'IFERROR', description: 'Returns a value if expression is an error' },

  // Lookup & Reference
  { name: 'VLOOKUP', description: 'Vertical lookup' },
  { name: 'HLOOKUP', description: 'Horizontal lookup' },
  { name: 'INDEX', description: 'Returns a value from a position' },
  { name: 'MATCH', description: 'Returns the position of a value' },
  { name: 'INDIRECT', description: 'Returns the reference specified by a string' },
  { name: 'CHOOSE', description: 'Chooses a value from a list' },
  { name: 'OFFSET', description: 'Returns a reference offset from a given reference' },
];

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
