/**
 * NameBox — the Excel "Name Box" for `<CasualSheets chrome>`.
 *
 * A small editable box that mirrors the active cell's A1 reference live, and
 * lets you navigate by typing a reference (e.g. `B5` or `A1:C3`) + Enter. Self-
 * contained: reads the active selection through `CasualSheetsAPI` and navigates
 * through the FUniver facade (`getRange(a1).activate()`) — no app context, no
 * named-range support (defined names land with the rich name-box later).
 *
 * The orchestrator uses this to replace the static name-box span currently
 * inside FormulaBar.tsx; this component does NOT edit FormulaBar.
 */

import { useEffect, useRef, useState, type CSSProperties, type KeyboardEvent } from 'react';
import { ICommandService } from '@univerjs/core';
import type { CasualSheetsAPI } from '../sheets/api';
import { Icon } from './Icon';

/** A defined name as surfaced by the FWorkbook facade. */
interface DefinedNameEntry {
  name: string;
  ref: string;
}

/**
 * Read the active workbook's defined names through the FWorkbook facade
 * (`getDefinedNames()` → `FDefinedName[]`, each with `getName()` +
 * `getFormulaOrRefString()`). Returns [] when the facade/API is unavailable —
 * never invents an API.
 */
function readDefinedNames(api: CasualSheetsAPI | null): DefinedNameEntry[] {
  if (!api) return [];
  try {
    const wb = api.univer.getActiveWorkbook() as unknown as {
      getDefinedNames?: () => Array<{
        getName?: () => string;
        getFormulaOrRefString?: () => string;
      }>;
    } | null;
    const names = wb?.getDefinedNames?.();
    if (!names) return [];
    return names
      .map((dn) => ({
        name: dn.getName?.() ?? '',
        ref: dn.getFormulaOrRefString?.() ?? '',
      }))
      .filter((e) => e.name !== '');
  } catch {
    return [];
  }
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

/** The active selection's top-left A1 reference, or '' when unavailable. */
function readActiveRef(api: CasualSheetsAPI): string {
  const sel = api.getSelection();
  if (!sel) return '';
  const { startRow, startColumn } = sel.range;
  return colToLetters(startColumn) + (startRow + 1);
}

const BOX_STYLE: CSSProperties = {
  flex: '0 0 auto',
  width: 80,
  height: 24,
  padding: '0 8px',
  border: '1px solid var(--cs-chrome-border, #e6e9ee)',
  borderRadius: 4,
  background: 'var(--cs-chrome-input-bg, #fff)',
  color: 'var(--cs-chrome-fg, #201f1e)',
  font: 'inherit',
  fontSize: 13,
  textAlign: 'center',
  boxSizing: 'border-box',
};

const ROOT_STYLE: CSSProperties = {
  flex: '0 0 auto',
  display: 'inline-flex',
  alignItems: 'center',
  position: 'relative',
};

const DROPDOWN_BTN_STYLE: CSSProperties = {
  flex: '0 0 auto',
  width: 20,
  height: 24,
  marginLeft: 2,
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

const MENU_STYLE: CSSProperties = {
  position: 'absolute',
  top: '100%',
  left: 0,
  marginTop: 4,
  zIndex: 1000,
  minWidth: 160,
  maxHeight: 280,
  overflowY: 'auto',
  padding: 4,
  borderRadius: 8,
  border: '1px solid var(--cs-chrome-border, #e6e9ee)',
  background: 'var(--cs-chrome-input-bg, #fff)',
  boxShadow: '0 6px 20px rgba(0,0,0,0.18)',
};

const MENU_ITEM_STYLE: CSSProperties = {
  display: 'flex',
  alignItems: 'baseline',
  justifyContent: 'space-between',
  gap: 12,
  width: '100%',
  height: 28,
  padding: '0 8px',
  border: 'none',
  borderRadius: 6,
  background: 'transparent',
  color: 'var(--cs-chrome-fg, #201f1e)',
  font: 'inherit',
  fontSize: 13,
  textAlign: 'left',
  cursor: 'pointer',
  boxSizing: 'border-box',
};

const MENU_EMPTY_STYLE: CSSProperties = {
  padding: '6px 8px',
  color: 'var(--cs-chrome-muted, #8a8886)',
  fontSize: 12,
  whiteSpace: 'nowrap',
};

const ITEM_REF_STYLE: CSSProperties = {
  color: 'var(--cs-chrome-muted, #8a8886)',
  fontSize: 11,
  flex: '0 0 auto',
};

export interface NameBoxProps {
  /** Live API, or `null` until the editor is ready. */
  api: CasualSheetsAPI | null;
}

export function NameBox({ api }: NameBoxProps) {
  // The A1 reference of the active cell.
  const [ref, setRef] = useState('');
  // null = mirror the active cell; a string = the user is typing.
  const [draft, setDraft] = useState<string | null>(null);
  const draftRef = useRef<string | null>(null);
  draftRef.current = draft;
  const inputRef = useRef<HTMLInputElement>(null);
  // The defined-names dropdown: open state, current entries, and a root ref for
  // outside-pointerdown detection (mirrors ColorPicker's popover pattern).
  const [open, setOpen] = useState(false);
  const [names, setNames] = useState<DefinedNameEntry[]>([]);
  const rootRef = useRef<HTMLDivElement>(null);

  // Reflect the active cell: subscribe to command activity (covers selection
  // moves) and re-read the reference — unless the user is mid-edit.
  useEffect(() => {
    if (!api) return;
    const refresh = () => {
      if (draftRef.current !== null) return;
      setRef(readActiveRef(api));
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

  // While the dropdown is open, close it on Escape or any outside pointerdown.
  useEffect(() => {
    if (!open) return;
    const onPointerDown = (e: PointerEvent) => {
      if (rootRef.current && !rootRef.current.contains(e.target as Node)) setOpen(false);
    };
    const onKeyDown = (e: globalThis.KeyboardEvent) => {
      if (e.key === 'Escape') setOpen(false);
    };
    document.addEventListener('pointerdown', onPointerDown, true);
    document.addEventListener('keydown', onKeyDown);
    return () => {
      document.removeEventListener('pointerdown', onPointerDown, true);
      document.removeEventListener('keydown', onKeyDown);
    };
  }, [open]);

  /** Navigate to the typed A1 reference / range. Invalid input is ignored. */
  const navigate = (text: string) => {
    const t = text.trim();
    setDraft(null);
    if (!api || t === '') {
      setRef(readActiveRef(api ?? ({} as CasualSheetsAPI)));
      return;
    }
    const sheet = api.univer.getActiveWorkbook()?.getActiveSheet();
    if (!sheet) return;
    try {
      // FRange.getRange parses an A1 string ("B5") or range ("A1:C3").
      const range = (sheet as unknown as { getRange(a1: string): { activate(): void } }).getRange(
        t,
      );
      range.activate();
    } catch {
      // Invalid reference — fall back to the current selection.
      setRef(readActiveRef(api));
    }
  };

  // Open the dropdown: refresh the defined-name list from the live workbook,
  // then toggle. mousedown (not click) so the grid selection isn't lost first.
  const toggleDropdown = () => {
    setOpen((cur) => {
      const next = !cur;
      if (next) setNames(readDefinedNames(api));
      return next;
    });
  };

  // Navigate to a defined name's range. The ref may carry a sheet prefix
  // ("Sheet1!$A$1") and absolute `$` markers — strip both, then reuse the same
  // getRange(...).activate() path the typed-reference navigation uses.
  const navigateName = (entry: DefinedNameEntry) => {
    setOpen(false);
    const raw = entry.ref.trim();
    if (raw === '') return;
    const afterSheet = raw.includes('!') ? raw.slice(raw.lastIndexOf('!') + 1) : raw;
    const a1 = afterSheet.replace(/\$/g, '');
    navigate(a1);
  };

  const onKeyDown = (e: KeyboardEvent<HTMLInputElement>) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      navigate((e.target as HTMLInputElement).value);
      (e.target as HTMLInputElement).blur();
    } else if (e.key === 'Escape') {
      e.preventDefault();
      setDraft(null);
      setRef(api ? readActiveRef(api) : '');
      (e.target as HTMLInputElement).blur();
    }
  };

  const shown = draft ?? ref;

  return (
    <div ref={rootRef} style={ROOT_STYLE}>
      <input
        ref={inputRef}
        type="text"
        aria-label="Name box"
        title="Name box"
        data-testid="cs-namebox-input"
        style={BOX_STYLE}
        value={shown}
        disabled={!api}
        onChange={(e) => setDraft(e.target.value)}
        onKeyDown={onKeyDown}
        onFocus={(e) => e.currentTarget.select()}
        onBlur={() => {
          // Abandon any uncommitted edit without navigating.
          if (draftRef.current !== null) {
            setDraft(null);
            setRef(api ? readActiveRef(api) : '');
          }
        }}
      />
      <button
        type="button"
        title="Defined names"
        aria-label="Defined names"
        aria-haspopup="true"
        aria-expanded={open}
        data-testid="cs-namebox-dropdown"
        style={{
          ...DROPDOWN_BTN_STYLE,
          background: open ? 'var(--cs-chrome-active, #e6f3f7)' : 'transparent',
          color: open ? 'var(--cs-chrome-active-fg, #0e7490)' : DROPDOWN_BTN_STYLE.color,
        }}
        disabled={!api}
        // mousedown (not click) so the grid's selection isn't lost first.
        onMouseDown={(e) => {
          e.preventDefault();
          toggleDropdown();
        }}
      >
        <Icon name="expand_more" size={18} />
      </button>
      {open && (
        <div style={MENU_STYLE} role="menu" data-testid="cs-namebox-menu">
          {names.length === 0 ? (
            <div style={MENU_EMPTY_STYLE}>No names</div>
          ) : (
            names.map((entry, i) => (
              <button
                key={entry.name}
                type="button"
                role="menuitem"
                title={entry.ref || entry.name}
                data-testid={`cs-namebox-name-${i}`}
                style={MENU_ITEM_STYLE}
                onMouseDown={(e) => {
                  e.preventDefault();
                  navigateName(entry);
                }}
                onMouseEnter={(e) => {
                  e.currentTarget.style.background = 'var(--cs-chrome-hover, rgba(0,0,0,0.06))';
                }}
                onMouseLeave={(e) => {
                  e.currentTarget.style.background = 'transparent';
                }}
              >
                <span>{entry.name}</span>
                {entry.ref && <span style={ITEM_REF_STYLE}>{entry.ref}</span>}
              </button>
            ))
          )}
        </div>
      )}
    </div>
  );
}
