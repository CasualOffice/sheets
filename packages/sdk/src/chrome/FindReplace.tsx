/**
 * FindReplace — a custom find/replace dialog for `<CasualSheets chrome>`.
 *
 * Univer's own find-replace dialog doesn't render in the SDK's headless mount
 * (it needs Univer's UI/overlay layer, which is dormant when header/toolbar/
 * footer are off), so the chrome ships its own — driven purely through the
 * facade: search reads the active sheet's `cellData` from `getSnapshot()`,
 * navigation `activate()`s the matching cell, and replace `setValue()`s it.
 *
 * Opens on Ctrl/Cmd+F (find) or Ctrl/Cmd+H (replace) while the editor has focus;
 * Escape closes. Self-managing — just mount it in the chrome with the live api.
 */

import {
  useEffect,
  useMemo,
  useRef,
  useState,
  type CSSProperties,
  type KeyboardEvent,
} from 'react';
import type { CasualSheetsAPI } from '../sheets/api';
import { Icon } from './Icon';

interface Match {
  row: number;
  col: number;
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type AnySnapshot = any;

/** Collect matches on the active sheet, in row-major order. Searches each
 *  cell's value (`v`) and formula (`f`) text. */
function findMatches(api: CasualSheetsAPI, query: string, matchCase: boolean): Match[] {
  if (!query) return [];
  const snap = api.getSnapshot() as AnySnapshot;
  const sheetId =
    api.getSelection()?.sheetId ?? api.univer.getActiveWorkbook()?.getActiveSheet()?.getSheetId();
  const sheet = sheetId ? snap?.sheets?.[sheetId] : undefined;
  const cellData = sheet?.cellData as Record<string, Record<string, AnySnapshot>> | undefined;
  if (!cellData) return [];
  const needle = matchCase ? query : query.toLowerCase();
  const out: Match[] = [];
  for (const rk of Object.keys(cellData).sort((a, b) => Number(a) - Number(b))) {
    const rowObj = cellData[rk];
    if (!rowObj) continue;
    for (const ck of Object.keys(rowObj).sort((a, b) => Number(a) - Number(b))) {
      const cell = rowObj[ck];
      if (!cell) continue;
      const hay = `${cell.v ?? ''}${cell.f ? ` ${cell.f}` : ''}`;
      const cmp = matchCase ? hay : hay.toLowerCase();
      if (cmp.includes(needle)) out.push({ row: Number(rk), col: Number(ck) });
    }
  }
  return out;
}

const WRAP_STYLE: CSSProperties = {
  position: 'absolute',
  top: 8,
  right: 12,
  zIndex: 1100,
  width: 320,
  padding: 10,
  borderRadius: 8,
  border: '1px solid var(--cs-chrome-border, #e6e9ee)',
  background: 'var(--cs-chrome-input-bg, #fff)',
  boxShadow: '0 8px 28px rgba(0,0,0,0.18)',
  font: 'inherit',
  fontSize: 13,
  color: 'var(--cs-chrome-fg, #201f1e)',
};

const ROW_STYLE: CSSProperties = { display: 'flex', alignItems: 'center', gap: 6, marginBottom: 6 };

const INPUT_STYLE: CSSProperties = {
  flex: '1 1 auto',
  height: 26,
  padding: '0 8px',
  border: '1px solid var(--cs-chrome-border, #cdd3db)',
  borderRadius: 4,
  background: 'var(--cs-chrome-input-bg, #fff)',
  color: 'var(--cs-chrome-fg, #201f1e)',
  font: 'inherit',
  fontSize: 13,
  boxSizing: 'border-box',
};

const BTN_STYLE: CSSProperties = {
  height: 26,
  padding: '0 10px',
  border: '1px solid var(--cs-chrome-border, #cdd3db)',
  borderRadius: 4,
  background: 'var(--cs-chrome-bg, #f3f5f8)',
  color: 'var(--cs-chrome-fg, #201f1e)',
  font: 'inherit',
  fontSize: 12,
  cursor: 'pointer',
};

const ICON_BTN_STYLE: CSSProperties = {
  width: 26,
  height: 26,
  display: 'inline-flex',
  alignItems: 'center',
  justifyContent: 'center',
  border: 'none',
  borderRadius: 4,
  background: 'transparent',
  color: 'var(--cs-chrome-fg, #201f1e)',
  cursor: 'pointer',
  padding: 0,
};

export interface FindReplaceProps {
  /** Live API, or `null` until the editor is ready. */
  api: CasualSheetsAPI | null;
  /**
   * Imperative open trigger from the dialog host. Bump this counter to open the
   * panel (e.g. when `openDialog('find-replace')` is called from a menu item).
   * The panel still self-opens on Ctrl/Cmd+F·H independently.
   */
  openSignal?: number;
  /** When the host opens it via `openSignal`, start in replace mode. */
  openInReplaceMode?: boolean;
}

export function FindReplace({ api, openSignal, openInReplaceMode }: FindReplaceProps) {
  const [open, setOpen] = useState(false);
  const [showReplace, setShowReplace] = useState(false);
  const [query, setQuery] = useState('');
  const [replaceText, setReplaceText] = useState('');
  const [matchCase, setMatchCase] = useState(false);
  const [idx, setIdx] = useState(0);
  const inputRef = useRef<HTMLInputElement>(null);

  // Open on Ctrl/Cmd+F (find) / Ctrl/Cmd+H (replace), close on Escape.
  useEffect(() => {
    const onKey = (e: globalThis.KeyboardEvent) => {
      const mod = e.metaKey || e.ctrlKey;
      if (mod && (e.key === 'f' || e.key === 'F')) {
        e.preventDefault();
        setShowReplace(false);
        setOpen(true);
        requestAnimationFrame(() => inputRef.current?.focus());
      } else if (mod && (e.key === 'h' || e.key === 'H')) {
        e.preventDefault();
        setShowReplace(true);
        setOpen(true);
        requestAnimationFrame(() => inputRef.current?.focus());
      }
    };
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  }, []);

  // Host-driven open (openDialog('find-replace')). `openSignal === undefined`
  // means "never opened by the host"; any defined value opens (and re-opening
  // bumps the counter). Skip the initial undefined→defined? No — only react when
  // a defined signal arrives.
  const lastSignal = useRef<number | undefined>(undefined);
  useEffect(() => {
    if (openSignal === undefined) return;
    if (lastSignal.current === openSignal) return;
    lastSignal.current = openSignal;
    setShowReplace(!!openInReplaceMode);
    setOpen(true);
    requestAnimationFrame(() => inputRef.current?.focus());
  }, [openSignal, openInReplaceMode]);

  const matches = useMemo(
    () => (api && open && query ? findMatches(api, query, matchCase) : []),
    [api, open, query, matchCase],
  );

  // Keep idx in range as matches change.
  useEffect(() => {
    setIdx((i) => (matches.length === 0 ? 0 : Math.min(i, matches.length - 1)));
  }, [matches.length]);

  if (!open || !api) return null;

  const go = (delta: number) => {
    if (matches.length === 0) return;
    const next = (idx + delta + matches.length) % matches.length;
    setIdx(next);
    activate(matches[next]);
  };

  const activate = (m: Match) => {
    const sheet = api.univer.getActiveWorkbook()?.getActiveSheet();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (sheet?.getRange(m.row, m.col) as any)?.activate?.();
  };

  const replaceOne = () => {
    if (matches.length === 0) return;
    const m = matches[Math.min(idx, matches.length - 1)];
    replaceAt(m);
    // Matches recompute via the query memo on next render; advance stays put.
  };

  const replaceAt = (m: Match) => {
    const sheet = api.univer.getActiveWorkbook()?.getActiveSheet();
    if (!sheet) return;
    const cur = String(sheet.getRange(m.row, m.col).getValue?.() ?? '');
    const next = replaceInString(cur, query, replaceText, matchCase);
    if (next !== cur) sheet.getRange(m.row, m.col).setValue(next);
  };

  const replaceAll = () => {
    const sheet = api.univer.getActiveWorkbook()?.getActiveSheet();
    if (!sheet) return;
    for (const m of matches) {
      const cur = String(sheet.getRange(m.row, m.col).getValue?.() ?? '');
      const next = replaceInString(cur, query, replaceText, matchCase);
      if (next !== cur) sheet.getRange(m.row, m.col).setValue(next);
    }
  };

  const onInputKey = (e: KeyboardEvent<HTMLInputElement>) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      go(e.shiftKey ? -1 : 1);
    } else if (e.key === 'Escape') {
      e.preventDefault();
      setOpen(false);
    }
  };

  const count = matches.length;
  const position = count === 0 ? 0 : Math.min(idx, count - 1) + 1;

  return (
    <div
      style={WRAP_STYLE}
      data-testid="cs-find-replace"
      role="dialog"
      aria-label="Find and replace"
    >
      <div style={ROW_STYLE}>
        <input
          ref={inputRef}
          type="text"
          aria-label="Find"
          placeholder="Find"
          data-testid="cs-find-input"
          style={INPUT_STYLE}
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          onKeyDown={onInputKey}
        />
        <span
          style={{ minWidth: 54, textAlign: 'center', color: 'var(--cs-chrome-muted, #6b7280)' }}
          data-testid="cs-find-count"
        >
          {position}/{count}
        </span>
        <button
          type="button"
          style={ICON_BTN_STYLE}
          aria-label="Previous match"
          data-testid="cs-find-prev"
          onMouseDown={(e) => {
            e.preventDefault();
            go(-1);
          }}
        >
          <Icon name="keyboard_arrow_up" size={18} />
        </button>
        <button
          type="button"
          style={ICON_BTN_STYLE}
          aria-label="Next match"
          data-testid="cs-find-next"
          onMouseDown={(e) => {
            e.preventDefault();
            go(1);
          }}
        >
          <Icon name="keyboard_arrow_down" size={18} />
        </button>
        <button
          type="button"
          style={ICON_BTN_STYLE}
          aria-label="Close"
          data-testid="cs-find-close"
          onMouseDown={(e) => {
            e.preventDefault();
            setOpen(false);
          }}
        >
          <Icon name="close" size={18} />
        </button>
      </div>

      {showReplace && (
        <div style={ROW_STYLE}>
          <input
            type="text"
            aria-label="Replace with"
            placeholder="Replace with"
            data-testid="cs-replace-input"
            style={INPUT_STYLE}
            value={replaceText}
            onChange={(e) => setReplaceText(e.target.value)}
          />
          <button
            type="button"
            style={BTN_STYLE}
            data-testid="cs-replace-one"
            onMouseDown={(e) => {
              e.preventDefault();
              replaceOne();
            }}
          >
            Replace
          </button>
          <button
            type="button"
            style={BTN_STYLE}
            data-testid="cs-replace-all"
            onMouseDown={(e) => {
              e.preventDefault();
              replaceAll();
            }}
          >
            All
          </button>
        </div>
      )}

      <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
        <label style={{ display: 'inline-flex', alignItems: 'center', gap: 4, cursor: 'pointer' }}>
          <input
            type="checkbox"
            data-testid="cs-find-match-case"
            checked={matchCase}
            onChange={(e) => setMatchCase(e.target.checked)}
          />
          Match case
        </label>
        {!showReplace && (
          <button
            type="button"
            style={{
              ...BTN_STYLE,
              marginLeft: 'auto',
              border: 'none',
              background: 'transparent',
              color: 'var(--cs-chrome-active-fg, #0e7490)',
            }}
            data-testid="cs-find-toggle-replace"
            onMouseDown={(e) => {
              e.preventDefault();
              setShowReplace(true);
            }}
          >
            Replace…
          </button>
        )}
      </div>
    </div>
  );
}

/** Replace all occurrences of `query` in `text` (case-(in)sensitive). */
function replaceInString(
  text: string,
  query: string,
  replacement: string,
  matchCase: boolean,
): string {
  if (!query) return text;
  if (matchCase) return text.split(query).join(replacement);
  // Case-insensitive: walk and rebuild preserving non-matched casing.
  const lower = text.toLowerCase();
  const needle = query.toLowerCase();
  let out = '';
  let i = 0;
  while (i < text.length) {
    if (lower.startsWith(needle, i)) {
      out += replacement;
      i += needle.length;
    } else {
      out += text[i];
      i += 1;
    }
  }
  return out;
}
