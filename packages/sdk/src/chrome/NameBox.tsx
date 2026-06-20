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
  );
}
