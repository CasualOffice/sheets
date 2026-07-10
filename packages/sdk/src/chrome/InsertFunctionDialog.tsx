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
 * InsertFunctionDialog — the SDK chrome's built-in Insert Function picker.
 *
 * A searchable, category-grouped list of common spreadsheet functions. Picking a
 * function seeds the active cell with `=NAME(` so the user can type arguments,
 * then focuses the grid. Applying is grounded in the real Sheets FUniver facade:
 *   - `api.univer.getActiveWorkbook()?.getActiveSheet()?.getActiveRange()` gives
 *     the live FRange (see FormatCellsDialog `activeRange`);
 *   - `FRange.setValue('=NAME(')` writes the formula stub into the range's
 *     top-left cell (verified in `@univerjs/sheets/lib/types/facade/f-range.d.ts`
 *     L814 `setValue(value: CellValue | ICellData): FRange`);
 *   - `FRange.getA1Notation()` (f-range.d.ts L1220) labels the target cell;
 *   - `api.focus()` returns keyboard focus to the grid for argument entry.
 *
 * Note on the caret: the installed FRange facade has no "enter cell edit mode
 * with the caret parked after the open-paren" method — `setValue` commits the
 * cell content and Univer re-parses it. So on choose we write `=NAME(` and focus
 * the grid; the user double-clicks / F2s to continue typing arguments. That's the
 * closest real, installed behaviour (recorded in the dialog's limitations).
 *
 * Mounted by `<DialogHost>` when `openDialog('insert-function')` is called and no
 * host override is registered.
 */

import { useMemo, useState, type CSSProperties } from 'react';
import type { DialogComponentProps } from './extensions';
import type { CasualSheetsAPI } from '../sheets/api';
import { Dialog } from './Dialog';
import {
  DIALOG_BTN_PRIMARY_STYLE,
  DIALOG_BTN_SECONDARY_STYLE,
  DIALOG_INPUT_STYLE,
} from './dialog-styles';

/** A single function entry in the picker. */
interface FunctionEntry {
  /** Canonical function name (upper-case), inserted verbatim. */
  name: string;
  /** Category bucket for grouping. */
  category: string;
  /** One-line description shown under the list / in the preview. */
  description: string;
  /** Argument-signature hint shown in the preview, e.g. `SUM(value1, [value2, …])`. */
  syntax: string;
}

/**
 * The common-function catalog. Deliberately a focused, Google-Sheets-style
 * "common functions" set rather than the full library — this is a quick-insert
 * picker, not a reference. Grouped by `category`.
 */
const FUNCTIONS: FunctionEntry[] = [
  // Math
  {
    name: 'SUM',
    category: 'Math',
    description: 'Sum of a set of numbers.',
    syntax: 'SUM(value1, [value2, …])',
  },
  {
    name: 'PRODUCT',
    category: 'Math',
    description: 'Product of a set of numbers.',
    syntax: 'PRODUCT(factor1, [factor2, …])',
  },
  {
    name: 'ROUND',
    category: 'Math',
    description: 'Round a number to a given number of places.',
    syntax: 'ROUND(value, [places])',
  },
  {
    name: 'ABS',
    category: 'Math',
    description: 'Absolute value of a number.',
    syntax: 'ABS(value)',
  },
  {
    name: 'MOD',
    category: 'Math',
    description: 'Remainder after division.',
    syntax: 'MOD(dividend, divisor)',
  },
  {
    name: 'POWER',
    category: 'Math',
    description: 'A number raised to a power.',
    syntax: 'POWER(base, exponent)',
  },
  // Statistical
  {
    name: 'AVERAGE',
    category: 'Statistical',
    description: 'Arithmetic mean of a set of numbers.',
    syntax: 'AVERAGE(value1, [value2, …])',
  },
  {
    name: 'COUNT',
    category: 'Statistical',
    description: 'Count of numeric values in a range.',
    syntax: 'COUNT(value1, [value2, …])',
  },
  {
    name: 'COUNTA',
    category: 'Statistical',
    description: 'Count of non-empty values in a range.',
    syntax: 'COUNTA(value1, [value2, …])',
  },
  {
    name: 'MAX',
    category: 'Statistical',
    description: 'Largest value in a set.',
    syntax: 'MAX(value1, [value2, …])',
  },
  {
    name: 'MIN',
    category: 'Statistical',
    description: 'Smallest value in a set.',
    syntax: 'MIN(value1, [value2, …])',
  },
  {
    name: 'MEDIAN',
    category: 'Statistical',
    description: 'Median of a set of numbers.',
    syntax: 'MEDIAN(value1, [value2, …])',
  },
  // Logical
  {
    name: 'IF',
    category: 'Logical',
    description: 'Return one value if a condition is true, another if false.',
    syntax: 'IF(condition, value_if_true, value_if_false)',
  },
  {
    name: 'IFS',
    category: 'Logical',
    description: 'Test multiple conditions, return the first match.',
    syntax: 'IFS(condition1, value1, [condition2, value2, …])',
  },
  {
    name: 'AND',
    category: 'Logical',
    description: 'True when all arguments are true.',
    syntax: 'AND(logical1, [logical2, …])',
  },
  {
    name: 'OR',
    category: 'Logical',
    description: 'True when any argument is true.',
    syntax: 'OR(logical1, [logical2, …])',
  },
  {
    name: 'IFERROR',
    category: 'Logical',
    description: 'Return a fallback when a formula errors.',
    syntax: 'IFERROR(value, value_if_error)',
  },
  // Lookup
  {
    name: 'VLOOKUP',
    category: 'Lookup',
    description: 'Search a column for a key, return a value from the same row.',
    syntax: 'VLOOKUP(search_key, range, index, [is_sorted])',
  },
  {
    name: 'HLOOKUP',
    category: 'Lookup',
    description: 'Search a row for a key, return a value from the same column.',
    syntax: 'HLOOKUP(search_key, range, index, [is_sorted])',
  },
  {
    name: 'INDEX',
    category: 'Lookup',
    description: 'Value at a given row/column offset in a range.',
    syntax: 'INDEX(reference, [row], [column])',
  },
  {
    name: 'MATCH',
    category: 'Lookup',
    description: 'Position of a value within a range.',
    syntax: 'MATCH(search_key, range, [search_type])',
  },
  {
    name: 'XLOOKUP',
    category: 'Lookup',
    description: 'Search a range and return a corresponding value.',
    syntax: 'XLOOKUP(search_key, lookup_range, result_range)',
  },
  // Text
  {
    name: 'CONCATENATE',
    category: 'Text',
    description: 'Join strings end to end.',
    syntax: 'CONCATENATE(string1, [string2, …])',
  },
  {
    name: 'LEFT',
    category: 'Text',
    description: 'Leftmost characters of a string.',
    syntax: 'LEFT(string, [num_chars])',
  },
  {
    name: 'RIGHT',
    category: 'Text',
    description: 'Rightmost characters of a string.',
    syntax: 'RIGHT(string, [num_chars])',
  },
  {
    name: 'MID',
    category: 'Text',
    description: 'Characters from the middle of a string.',
    syntax: 'MID(string, start, length)',
  },
  { name: 'LEN', category: 'Text', description: 'Length of a string.', syntax: 'LEN(string)' },
  {
    name: 'TRIM',
    category: 'Text',
    description: 'Remove leading/trailing/duplicate spaces.',
    syntax: 'TRIM(string)',
  },
  // Date
  { name: 'TODAY', category: 'Date', description: "Today's date.", syntax: 'TODAY()' },
  { name: 'NOW', category: 'Date', description: 'Current date and time.', syntax: 'NOW()' },
  {
    name: 'DATE',
    category: 'Date',
    description: 'Build a date from year, month, day.',
    syntax: 'DATE(year, month, day)',
  },
  {
    name: 'DATEDIF',
    category: 'Date',
    description: 'Difference between two dates in a chosen unit.',
    syntax: 'DATEDIF(start_date, end_date, unit)',
  },
];

/** Ordered category list, so groups render in a stable, sensible order. */
const CATEGORY_ORDER = ['Math', 'Statistical', 'Logical', 'Lookup', 'Text', 'Date'];

/** The active FRange, or null when there is no selection. */
function activeRange(api: CasualSheetsAPI) {
  return api.univer.getActiveWorkbook()?.getActiveSheet()?.getActiveRange() ?? null;
}

/**
 * Insert `=NAME(` into the top-left cell of the active range via the facade,
 * then focus the grid for argument entry. Returns false when there's no
 * selection to write into.
 */
function insertFunction(api: CasualSheetsAPI, fn: FunctionEntry): boolean {
  const range = activeRange(api);
  if (!range) return false;

  // Zero-arg functions (TODAY / NOW) are complete on their own; for the rest,
  // leave the open-paren so the user continues typing arguments.
  const zeroArg = fn.syntax.endsWith('()');
  const formula = zeroArg ? `=${fn.name}()` : `=${fn.name}(`;

  // FRange.setValue writes into the range's top-left cell (f-range.d.ts L814).
  range.setValue(formula);

  // Return keyboard focus to the grid so the user can keep typing arguments.
  api.focus();
  return true;
}

const SEARCH_STYLE: CSSProperties = {
  ...DIALOG_INPUT_STYLE,
  width: '100%',
  marginBottom: 12,
};

const RANGE_NOTE_STYLE: CSSProperties = {
  fontSize: 12,
  color: 'var(--cs-chrome-muted, #605e5c)',
  marginBottom: 12,
};

const LIST_STYLE: CSSProperties = {
  maxHeight: 280,
  overflow: 'auto',
  border: '1px solid var(--cs-chrome-border, #edeff3)',
  borderRadius: 8,
};

const GROUP_HEADER_STYLE: CSSProperties = {
  position: 'sticky',
  top: 0,
  padding: '6px 12px',
  fontSize: 11,
  fontWeight: 600,
  letterSpacing: 0.4,
  textTransform: 'uppercase',
  color: 'var(--cs-chrome-muted, #605e5c)',
  background: 'var(--cs-chrome-input-bg, #fff)',
  borderBottom: '1px solid var(--cs-chrome-border, #edeff3)',
};

function itemStyle(selected: boolean): CSSProperties {
  return {
    display: 'block',
    width: '100%',
    textAlign: 'left',
    padding: '8px 12px',
    border: 'none',
    borderBottom: '1px solid var(--cs-chrome-border, #f3f4f6)',
    background: selected ? 'var(--cs-chrome-hover-bg, #eef6f8)' : 'transparent',
    color: 'var(--cs-chrome-fg, #201f1e)',
    font: 'inherit',
    fontSize: 13,
    cursor: 'pointer',
  };
}

const ITEM_NAME_STYLE: CSSProperties = {
  fontWeight: 600,
  fontFamily: 'var(--cs-chrome-mono, ui-monospace, SFMono-Regular, Menlo, monospace)',
};

const ITEM_DESC_STYLE: CSSProperties = {
  fontSize: 12,
  color: 'var(--cs-chrome-muted, #605e5c)',
  marginTop: 2,
};

const PREVIEW_STYLE: CSSProperties = {
  marginTop: 12,
  padding: '10px 12px',
  borderRadius: 8,
  background: 'var(--cs-chrome-subtle-bg, #f6f8fa)',
  fontSize: 12,
  color: 'var(--cs-chrome-fg, #201f1e)',
};

const PREVIEW_SYNTAX_STYLE: CSSProperties = {
  fontFamily: 'var(--cs-chrome-mono, ui-monospace, SFMono-Regular, Menlo, monospace)',
  fontWeight: 600,
  marginBottom: 4,
};

export function InsertFunctionDialog({ api, onClose }: DialogComponentProps) {
  const [query, setQuery] = useState('');
  const [selectedName, setSelectedName] = useState<string | null>(null);

  // Label the target cell once for the header hint.
  const rangeLabel = useMemo(() => {
    const fRange = activeRange(api) as unknown as { getA1Notation?: () => string } | null;
    return fRange?.getA1Notation?.() ?? null;
  }, [api]);

  const hasSelection = activeRange(api) !== null;

  // Filter by name / description / category, case-insensitive.
  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) return FUNCTIONS;
    return FUNCTIONS.filter(
      (fn) =>
        fn.name.toLowerCase().includes(q) ||
        fn.description.toLowerCase().includes(q) ||
        fn.category.toLowerCase().includes(q),
    );
  }, [query]);

  // Group the filtered set by category, preserving CATEGORY_ORDER.
  const grouped = useMemo(() => {
    const byCat = new Map<string, FunctionEntry[]>();
    for (const fn of filtered) {
      const list = byCat.get(fn.category);
      if (list) list.push(fn);
      else byCat.set(fn.category, [fn]);
    }
    return CATEGORY_ORDER.filter((c) => byCat.has(c)).map((c) => ({
      category: c,
      items: byCat.get(c)!,
    }));
  }, [filtered]);

  const selected = useMemo(
    () => FUNCTIONS.find((fn) => fn.name === selectedName) ?? null,
    [selectedName],
  );

  const choose = (fn: FunctionEntry) => {
    if (insertFunction(api, fn)) onClose();
  };

  const insertSelected = () => {
    if (selected && insertFunction(api, selected)) onClose();
  };

  return (
    <Dialog
      title="Insert function"
      onClose={onClose}
      width={460}
      data-testid="cs-insert-function-dialog"
      footer={
        <>
          <button type="button" style={DIALOG_BTN_SECONDARY_STYLE} onClick={onClose}>
            Cancel
          </button>
          <button
            type="button"
            style={DIALOG_BTN_PRIMARY_STYLE}
            data-testid="cs-insert-function-insert"
            disabled={!hasSelection || !selected}
            onClick={insertSelected}
          >
            Insert
          </button>
        </>
      }
    >
      {hasSelection ? (
        <div style={RANGE_NOTE_STYLE} data-testid="cs-insert-function-range">
          Inserts into <strong>{rangeLabel ?? 'the active cell'}</strong>
        </div>
      ) : (
        <div style={RANGE_NOTE_STYLE} data-testid="cs-insert-function-no-selection">
          Select a cell first, then reopen this dialog.
        </div>
      )}

      <input
        style={SEARCH_STYLE}
        data-testid="cs-insert-function-search"
        type="search"
        placeholder="Search functions (name, category, or description)"
        value={query}
        autoFocus
        onChange={(e) => setQuery(e.target.value)}
      />

      <div
        style={LIST_STYLE}
        role="listbox"
        aria-label="Functions"
        data-testid="cs-insert-function-list"
      >
        {grouped.length === 0 && (
          <div style={{ padding: '12px', fontSize: 13, color: 'var(--cs-chrome-muted, #605e5c)' }}>
            No functions match &ldquo;{query}&rdquo;.
          </div>
        )}
        {grouped.map((group) => (
          <div key={group.category}>
            <div style={GROUP_HEADER_STYLE}>{group.category}</div>
            {group.items.map((fn) => (
              <button
                key={fn.name}
                type="button"
                role="option"
                aria-selected={selectedName === fn.name}
                data-testid={`cs-insert-function-item-${fn.name}`}
                style={itemStyle(selectedName === fn.name)}
                onClick={() => setSelectedName(fn.name)}
                onDoubleClick={() => choose(fn)}
              >
                <div style={ITEM_NAME_STYLE}>{fn.name}</div>
                <div style={ITEM_DESC_STYLE}>{fn.description}</div>
              </button>
            ))}
          </div>
        ))}
      </div>

      {selected && (
        <div style={PREVIEW_STYLE} data-testid="cs-insert-function-preview">
          <div style={PREVIEW_SYNTAX_STYLE}>{selected.syntax}</div>
          <div>{selected.description}</div>
        </div>
      )}
    </Dialog>
  );
}
