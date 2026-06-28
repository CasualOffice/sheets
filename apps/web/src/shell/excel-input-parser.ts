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
 * Parse a string the user typed into a cell the way Excel does:
 * recognise currency symbols, percent suffix, accounting parentheses,
 * thousands grouping, and signed numbers, then return the coerced
 * number. When the input is clearly not numeric, returns `unchanged`
 * so the caller leaves the cell as text.
 *
 * Pure function, no Univer dependencies — tested in isolation.
 *
 * Why this exists: Univer's default in-cell editor stores `$2,000` as
 * a text cell with `CellValueType.STRING`, which then trips the
 * "Number stored as text" alert. Excel parses `$2,000` to `2000` and
 * leaves the cell's number-format pattern to render the `$` back on
 * display. We mirror that.
 */
export type ExcelParseResult =
  | { kind: 'unchanged' }
  | {
      kind: 'number';
      value: number;
      hint: 'plain' | 'currency' | 'percent' | 'accounting' | 'currency-accounting';
      currencySymbol?: string;
    };

// Single-char currency symbols. Excel recognises most of these; we add
// the common Latin-American + Asian + African symbols that show up in
// real-world workbooks.
//
// Ordered by Unicode codepoint for readability — order has no semantic
// effect because we match each as a single-char prefix.
const CURRENCY_CHARS = new Set([
  '$', // dollar / peso (many locales)
  '¢', // cent
  '£', // pound sterling
  '¤', // generic currency
  '¥', // yen / yuan
  '₠', // ECU (legacy)
  '₡', // colón
  '₢', // cruzeiro
  '₣', // French franc (legacy)
  '₤', // lira (legacy)
  '₦', // naira
  '₩', // won
  '₪', // shekel
  '₫', // dong
  '€', // euro
  '₭', // kip
  '₮', // tugrik
  '₱', // peso (PH)
  '₲', // guarani
  '₹', // indian rupee
  '₴', // hryvnia
  '₵', // cedi
  '₸', // tenge
  '₺', // turkish lira
  '₼', // azerbaijan manat
  '₽', // ruble
  '₾', // georgian lari
  '₿', // bitcoin
  '﷼', // saudi riyal
  '฿', // baht
]);

// Multi-character currency prefixes / suffixes Excel routinely accepts.
// Order matters when matching — longer first so `NT$` beats `$`,
// `US$` beats `$`, `Mex$` beats `$`, etc.
const MULTI_CHAR_CURRENCIES = [
  'Mex$',
  'NT$',
  'HK$',
  'US$',
  'CA$',
  'AU$',
  'NZ$',
  'SG$',
  'R$',
  'S$',
  'CHF',
  'CFA',
  'CFP',
  'kr.', // some Nordic locales include the dot
  'zł',
  'Kč',
  'Ft',
  'Rp',
  'RM',
  'kr',
  'Fr',
  'lei',
  'Bs',
  'Br',
];

/**
 * Try every known currency prefix at `s[0..]`. Returns `null` if none
 * match. Walks longest-first so `NT$` doesn't get partially consumed
 * as `N` + `T$` (it wouldn't anyway, but be explicit).
 */
function stripCurrencyPrefix(s: string): { rest: string; symbol: string } | null {
  for (const sym of MULTI_CHAR_CURRENCIES) {
    if (s.startsWith(sym)) return { rest: s.slice(sym.length), symbol: sym };
  }
  if (s.length > 0 && CURRENCY_CHARS.has(s[0]!)) {
    return { rest: s.slice(1), symbol: s[0]! };
  }
  return null;
}

/**
 * Same idea, but as a suffix. Some locales (de-DE, sv-SE) write the
 * currency after the amount (`1.234,50 €`, `42 kr`). We support the
 * common suffix shape `<number><space?><symbol>` in addition to the
 * en-US prefix shape.
 */
function stripCurrencySuffix(s: string): { rest: string; symbol: string } | null {
  for (const sym of MULTI_CHAR_CURRENCIES) {
    if (s.endsWith(sym)) {
      const rest = s.slice(0, -sym.length).trimEnd();
      return { rest, symbol: sym };
    }
  }
  if (s.length > 0 && CURRENCY_CHARS.has(s[s.length - 1]!)) {
    return { rest: s.slice(0, -1).trimEnd(), symbol: s[s.length - 1]! };
  }
  return null;
}

/**
 * Validate that comma usage in the numeric body matches Excel's
 * thousands grouping (`1,234` `1,234,567` `12,345.67`). Bare `1,5` is
 * NOT valid en-US thousands — it's the de-DE decimal which we don't
 * support. We require commas to appear in groups-of-three or not at
 * all; this guards us from silently turning `1,5` into `15`.
 */
function isValidThousandsGrouping(body: string): boolean {
  if (!body.includes(',')) return true;
  // Allow leading sign, optional decimal tail. Integer part must be
  // 1–3 digits followed by repeated `,\d{3}` groups.
  return /^[+-]?\d{1,3}(,\d{3})+(\.\d+)?$/.test(body);
}

/**
 * The core parser. Returns a discriminated union so callers know
 * whether to mutate the cell or leave it alone.
 *
 * Input invariant: `raw` is whatever the editor handed us (un-trimmed).
 */
export function parseExcelStyleValue(raw: string): ExcelParseResult {
  if (typeof raw !== 'string') return { kind: 'unchanged' };
  const s0 = raw.trim();
  if (s0 === '') return { kind: 'unchanged' };

  // Force-text marker: leading apostrophe means "always text" — Excel
  // and Univer both respect this. We don't touch it.
  if (s0.startsWith("'")) return { kind: 'unchanged' };

  // Formula — out of scope; Univer's own pipeline handles `=...`.
  if (s0.startsWith('=')) return { kind: 'unchanged' };

  let body = s0;
  let negate = false;
  let isAccounting = false;
  let isPercent = false;
  let symbol: string | undefined;

  // 1. Accounting parens: `(1234)` → negative 1234. Recognise both
  //    `(1234)` and `($1234)` shapes — the inner side can still hold
  //    a currency symbol.
  if (body.startsWith('(') && body.endsWith(')') && body.length >= 3) {
    body = body.slice(1, -1).trim();
    negate = true;
    isAccounting = true;
  }

  // 2. Leading sign. Excel accepts `+500` and `-500`. We capture the
  //    sign here so a downstream currency / paren combo composes
  //    cleanly with it.
  if (body.startsWith('+')) {
    body = body.slice(1).trimStart();
  } else if (body.startsWith('-')) {
    body = body.slice(1).trimStart();
    negate = !negate;
  }

  // 3. Currency prefix. After stripping, allow another optional sign
  //    (`$-100` is real, Excel accepts it).
  const pref = stripCurrencyPrefix(body);
  if (pref) {
    body = pref.rest.trimStart();
    symbol = pref.symbol;
    if (body.startsWith('+')) body = body.slice(1).trimStart();
    else if (body.startsWith('-')) {
      body = body.slice(1).trimStart();
      negate = !negate;
    }
  }

  // 4. Trailing percent. Must come BEFORE trailing-currency detection
  //    so `12.5%` doesn't get mis-stripped as currency.
  if (body.endsWith('%')) {
    body = body.slice(0, -1).trimEnd();
    isPercent = true;
  }

  // 5. Trailing currency (suffix locales) — only if we didn't already
  //    consume a prefix. Avoids double-counting in mixed inputs.
  if (!symbol) {
    const suf = stripCurrencySuffix(body);
    if (suf) {
      body = suf.rest;
      symbol = suf.symbol;
    }
  }

  // 6. Body must now look like a number with optional thousands
  //    grouping. Reject `1,5` (de-DE decimal) — we're en-US only.
  if (!isValidThousandsGrouping(body)) return { kind: 'unchanged' };

  // Strip the now-validated thousands separators before Number().
  const numeric = body.replace(/,/g, '');

  // Reject inputs Number() would happily turn into NaN/Infinity, and
  // inputs that are just an empty string after stripping (e.g. the
  // user typed `$` alone — leave it as text).
  if (numeric === '' || numeric === '.' || numeric === '-' || numeric === '+') {
    return { kind: 'unchanged' };
  }
  const n = Number(numeric);
  if (!Number.isFinite(n)) return { kind: 'unchanged' };

  let value = negate ? -n : n;
  if (isPercent) value = value / 100;
  // Normalise -0 to 0 so downstream JSON / display doesn't surface it.
  if (value === 0) value = 0;

  // Currency + accounting compose — emit a combined hint so the caller
  // can pick `_("$"* #,##0_)` over plain `0_);(0)`.
  type Hint = 'plain' | 'currency' | 'percent' | 'accounting' | 'currency-accounting';
  let hint: Hint;
  if (isPercent) hint = 'percent';
  else if (symbol && isAccounting) hint = 'currency-accounting';
  else if (symbol) hint = 'currency';
  else if (isAccounting) hint = 'accounting';
  else hint = 'plain';

  return { kind: 'number', value, hint, ...(symbol ? { currencySymbol: symbol } : {}) };
}

/**
 * Build a number-format pattern for a hint + symbol. Used when the
 * destination cell has no existing format and we want to surface the
 * formatting Excel would apply on first entry.
 *
 * Returns `null` for the plain hint — plain numbers don't need a
 * pattern (cells default to General).
 */
export function suggestedPatternFor(
  hint: 'plain' | 'currency' | 'percent' | 'accounting' | 'currency-accounting',
  symbol: string | undefined,
): string | null {
  switch (hint) {
    case 'plain':
      return null;
    case 'percent':
      return '0.00%';
    case 'currency': {
      const s = symbol ?? '$';
      return `"${s}"#,##0.00`;
    }
    case 'accounting':
      return '#,##0.00_);(#,##0.00)';
    case 'currency-accounting': {
      const s = symbol ?? '$';
      // Excel's canonical accounting pattern. The `_(` and `_)` add a
      // trailing-space the width of a paren so positives align with
      // parenthesised negatives. The `"-"??` is the zero placeholder.
      return `_("${s}"* #,##0.00_);_("${s}"* (#,##0.00);_("${s}"* "-"??_);_(@_)`;
    }
  }
}
