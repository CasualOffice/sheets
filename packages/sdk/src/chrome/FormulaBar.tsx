/**
 * FormulaBar — minimal built-in formula bar for `<CasualSheets chrome>`.
 *
 * Second slice of the chrome lift (after the Toolbar). Self-contained: it reads
 * the active cell through `CasualSheetsAPI` and commits edits through the facade
 * — no app context, no formula autocomplete / name-box / insert-function dialog
 * (those land when the rich `apps/web` FormulaBar is lifted behind `chrome="full"`).
 *
 * Shows the active cell's A1 reference + its formula (or value), and lets you
 * edit it: `=…` commits as a formula, a number commits as a number, anything
 * else as text.
 */

import { useEffect, useRef, useState, type CSSProperties, type KeyboardEvent } from 'react';
import { ICommandService } from '@univerjs/core';
import type { CasualSheetsAPI } from '../sheets/api';
import { NameBox } from './NameBox';

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
  flex: '1 1 auto',
  minWidth: 0,
  height: 24,
  padding: '0 8px',
  border: '1px solid var(--cs-chrome-border, rgba(0,0,0,0.18))',
  borderRadius: 4,
  background: 'var(--cs-chrome-input-bg, #fff)',
  color: 'var(--cs-chrome-fg, #1f2329)',
  font: 'inherit',
  fontSize: 13,
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

  const commit = (text: string) => {
    setDraft(null);
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

  const onKeyDown = (e: KeyboardEvent<HTMLInputElement>) => {
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
      <input
        type="text"
        aria-label="Formula bar"
        data-testid="casual-sheets-formula-input"
        style={INPUT_STYLE}
        value={shown}
        onChange={(e) => setDraft(e.target.value)}
        onKeyDown={onKeyDown}
        onBlur={(e) => {
          if (draftRef.current !== null) commit(e.target.value);
        }}
        disabled={!api}
      />
    </div>
  );
}
