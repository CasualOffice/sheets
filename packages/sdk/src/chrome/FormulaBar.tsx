/**
 * FormulaBar — minimal built-in formula bar for `<CasualSheets chrome>`.
 *
 * The formula bar: a NameBox (active-cell ref + go-to navigation) and an editable
 * input that shows the active cell's formula (or value) and commits edits through
 * the facade — `=…` as a formula, a number as a number, else text. Typing a
 * function name after `=` shows an autocomplete dropdown (arrows + Enter/Tab to
 * complete, Escape to dismiss). Self-contained: drives the editor only via
 * `CasualSheetsAPI`.
 */

import {
  useEffect,
  useMemo,
  useRef,
  useState,
  type CSSProperties,
  type KeyboardEvent,
} from 'react';
import { ICommandService } from '@univerjs/core';
import type { CasualSheetsAPI } from '../sheets/api';
import { NameBox } from './NameBox';

// Curated common spreadsheet functions for the autocomplete dropdown. (Univer
// ships ~530; this covers what people actually type without bundling them all.)
const FUNCTIONS = [
  'ABS',
  'AND',
  'AVERAGE',
  'AVERAGEIF',
  'AVERAGEIFS',
  'CEILING',
  'CHOOSE',
  'COLUMN',
  'COLUMNS',
  'CONCAT',
  'CONCATENATE',
  'COUNT',
  'COUNTA',
  'COUNTBLANK',
  'COUNTIF',
  'COUNTIFS',
  'DATE',
  'DATEDIF',
  'DAY',
  'EDATE',
  'EOMONTH',
  'EXACT',
  'FILTER',
  'FIND',
  'FLOOR',
  'HLOOKUP',
  'HOUR',
  'IF',
  'IFERROR',
  'IFNA',
  'IFS',
  'INDEX',
  'INDIRECT',
  'INT',
  'ISBLANK',
  'ISERROR',
  'ISNA',
  'ISNUMBER',
  'ISTEXT',
  'LEFT',
  'LEN',
  'LOOKUP',
  'LOWER',
  'MATCH',
  'MAX',
  'MAXIFS',
  'MEDIAN',
  'MID',
  'MIN',
  'MINIFS',
  'MINUTE',
  'MOD',
  'MONTH',
  'NETWORKDAYS',
  'NOT',
  'NOW',
  'OFFSET',
  'OR',
  'POWER',
  'PRODUCT',
  'PROPER',
  'RAND',
  'RANDBETWEEN',
  'RANK',
  'REPLACE',
  'REPT',
  'RIGHT',
  'ROUND',
  'ROUNDDOWN',
  'ROUNDUP',
  'ROW',
  'ROWS',
  'SEARCH',
  'SECOND',
  'SORT',
  'SQRT',
  'SUBSTITUTE',
  'SUM',
  'SUMIF',
  'SUMIFS',
  'SUMPRODUCT',
  'SWITCH',
  'TEXT',
  'TEXTJOIN',
  'TODAY',
  'TRANSPOSE',
  'TRIM',
  'UNIQUE',
  'UPPER',
  'VALUE',
  'VLOOKUP',
  'WEEKDAY',
  'WORKDAY',
  'XLOOKUP',
  'XMATCH',
  'YEAR',
];

const TOKEN_RE = /([A-Za-z][A-Za-z0-9._]*)$/;

/** Function suggestions for the trailing identifier of a formula draft. */
function suggestFor(draft: string | null): string[] {
  if (!draft || !draft.startsWith('=')) return [];
  const m = TOKEN_RE.exec(draft);
  if (!m) return [];
  const prefix = m[1].toUpperCase();
  return FUNCTIONS.filter((f) => f.startsWith(prefix) && f !== prefix).slice(0, 8);
}

/** A1 column letters from a 0-based column index (0→A, 26→AA). */
function colToLetters(col: number): string {
  let s = '';
  let n = col;
  do {
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return s;
}

interface ActiveCell {
  a1: string;
  /** Formula text (`=…`) if the cell has one, else the raw value as a string. */
  text: string;
}

function readActiveCell(api: CasualSheetsAPI): ActiveCell | null {
  const sel = api.getSelection();
  const sheet = api.univer.getActiveWorkbook()?.getActiveSheet();
  if (!sel || !sheet) return null;
  const { startRow, startColumn } = sel.range;
  const range = sheet.getRange(startRow, startColumn);
  const formula = range.getFormula?.() || '';
  const raw = range.getValue?.();
  return {
    a1: colToLetters(startColumn) + (startRow + 1),
    text: formula || (raw == null ? '' : String(raw)),
  };
}

const BAR_STYLE: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 6,
  padding: '3px 6px',
  borderBottom: '1px solid var(--cs-chrome-border, rgba(0,0,0,0.12))',
  background: 'var(--cs-chrome-bg, #f8f9fa)',
  flex: '0 0 auto',
  font: 'inherit',
  fontSize: 13,
};

const FX_STYLE: CSSProperties = {
  flex: '0 0 auto',
  fontStyle: 'italic',
  color: 'var(--cs-chrome-muted, #6b7280)',
};

const INPUT_STYLE: CSSProperties = {
  width: '100%',
  height: 24,
  padding: '0 8px',
  border: '1px solid var(--cs-chrome-border, rgba(0,0,0,0.18))',
  borderRadius: 4,
  background: 'var(--cs-chrome-input-bg, #fff)',
  color: 'var(--cs-chrome-fg, #1f2329)',
  font: 'inherit',
  fontSize: 13,
  boxSizing: 'border-box',
};

const INPUT_WRAP_STYLE: CSSProperties = { position: 'relative', flex: '1 1 auto', minWidth: 0 };

const AC_STYLE: CSSProperties = {
  position: 'absolute',
  top: '100%',
  left: 0,
  marginTop: 2,
  minWidth: 200,
  maxHeight: 240,
  overflowY: 'auto',
  zIndex: 1000,
  padding: 4,
  border: '1px solid var(--cs-chrome-border, #e6e9ee)',
  borderRadius: 8,
  background: 'var(--cs-chrome-input-bg, #fff)',
  boxShadow: '0 6px 20px rgba(0,0,0,0.16)',
};

const AC_ITEM_STYLE: CSSProperties = {
  display: 'block',
  width: '100%',
  textAlign: 'left',
  height: 26,
  padding: '0 8px',
  border: 'none',
  borderRadius: 6,
  background: 'transparent',
  color: 'var(--cs-chrome-fg, #201f1e)',
  font: 'inherit',
  fontSize: 13,
  cursor: 'pointer',
};

export interface FormulaBarProps {
  /** Live API, or `null` until the editor is ready. */
  api: CasualSheetsAPI | null;
}

export function FormulaBar({ api }: FormulaBarProps) {
  const [cell, setCell] = useState<ActiveCell | null>(null);
  // null = mirror the active cell; a string = the user is editing.
  const [draft, setDraft] = useState<string | null>(null);
  const draftRef = useRef<string | null>(null);
  draftRef.current = draft;
  // Autocomplete: active suggestion index + a dismissed flag (Escape closes the
  // dropdown without cancelling the edit).
  const [acIdx, setAcIdx] = useState(0);
  const [acDismissed, setAcDismissed] = useState(false);

  useEffect(() => {
    if (!api) return;
    const refresh = () => {
      // Don't clobber an in-progress edit on unrelated command activity.
      if (draftRef.current !== null) return;
      setCell(readActiveCell(api));
    };
    refresh();
    const injector = (api.univer as unknown as { _injector?: { get(t: unknown): unknown } })
      ._injector;
    const cmd = injector?.get(ICommandService) as
      | { onCommandExecuted: (cb: () => void) => { dispose: () => void } }
      | undefined;
    const sub = cmd?.onCommandExecuted(() => refresh());
    return () => sub?.dispose();
  }, [api]);

  const suggestions = useMemo(() => suggestFor(draft), [draft]);
  const acOpen = !acDismissed && suggestions.length > 0;

  const commit = (text: string) => {
    setDraft(null);
    setAcDismissed(false);
    if (!api) return;
    const sel = api.getSelection();
    const sheet = api.univer.getActiveWorkbook()?.getActiveSheet();
    if (!sel || !sheet) return;
    const range = sheet.getRange(sel.range.startRow, sel.range.startColumn);
    const t = text.trim();
    if (t.startsWith('=')) range.setValue({ f: t });
    else if (t !== '' && !Number.isNaN(Number(t))) range.setValue(Number(t));
    else range.setValue(t);
  };

  const onType = (v: string) => {
    setDraft(v);
    setAcIdx(0);
    setAcDismissed(false);
  };

  // Replace the trailing identifier with the chosen function + '('.
  const complete = (fn: string) => {
    const base = draft ?? '';
    setDraft(base.replace(TOKEN_RE, `${fn}(`));
    setAcIdx(0);
  };

  const onKeyDown = (e: KeyboardEvent<HTMLInputElement>) => {
    if (acOpen) {
      if (e.key === 'ArrowDown') {
        e.preventDefault();
        setAcIdx((i) => Math.min(i + 1, suggestions.length - 1));
        return;
      }
      if (e.key === 'ArrowUp') {
        e.preventDefault();
        setAcIdx((i) => Math.max(i - 1, 0));
        return;
      }
      if (e.key === 'Enter' || e.key === 'Tab') {
        e.preventDefault();
        complete(suggestions[acIdx] ?? suggestions[0]);
        return;
      }
      if (e.key === 'Escape') {
        e.preventDefault();
        setAcDismissed(true);
        return;
      }
    }
    if (e.key === 'Enter') {
      e.preventDefault();
      commit((e.target as HTMLInputElement).value);
    } else if (e.key === 'Escape') {
      e.preventDefault();
      setDraft(null);
      (e.target as HTMLInputElement).blur();
    }
  };

  const shown = draft ?? cell?.text ?? '';

  return (
    <div style={BAR_STYLE} data-testid="casual-sheets-formula-bar">
      <NameBox api={api} />
      <span style={FX_STYLE} aria-hidden>
        fx
      </span>
      <div style={INPUT_WRAP_STYLE}>
        <input
          type="text"
          aria-label="Formula bar"
          data-testid="casual-sheets-formula-input"
          style={INPUT_STYLE}
          value={shown}
          onChange={(e) => onType(e.target.value)}
          onKeyDown={onKeyDown}
          onBlur={(e) => {
            if (draftRef.current !== null) commit(e.target.value);
          }}
          disabled={!api}
        />
        {acOpen && (
          <div style={AC_STYLE} data-testid="cs-formula-suggestions" role="listbox">
            {suggestions.map((fn, i) => (
              <button
                key={fn}
                type="button"
                role="option"
                aria-selected={i === acIdx}
                data-testid={`cs-formula-suggestion-${fn}`}
                style={{
                  ...AC_ITEM_STYLE,
                  background: i === acIdx ? 'var(--cs-chrome-active, #e6f3f7)' : 'transparent',
                  color: i === acIdx ? 'var(--cs-chrome-active-fg, #0e7490)' : AC_ITEM_STYLE.color,
                }}
                // mousedown+preventDefault so the input keeps focus (no blur-commit).
                onMouseDown={(e) => {
                  e.preventDefault();
                  complete(fn);
                }}
                onMouseEnter={() => setAcIdx(i)}
              >
                {fn}
              </button>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}
