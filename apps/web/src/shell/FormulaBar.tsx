import {
  useCallback,
  useEffect,
  useLayoutEffect,
  useRef,
  useState,
  type ChangeEvent,
  type KeyboardEvent,
} from 'react';
import { flushSync } from 'react-dom';
import { useUniverAPI } from '../use-univer';
import { useActiveCellState } from '../hooks/useActiveCellState';
import {
  canInsertRefAtCaret,
  commitToActiveCell,
  cycleAbsoluteRefAtCaret,
  insertRefAtCaret,
  quoteSheetName,
  type CommitDirection,
} from './cell-actions';
import { Icon } from './Icon';
import { Tooltip } from './Tooltip';
import {
  extractFunctionFragment,
  suggestFunctions,
  type FormulaFn,
} from './formula-functions';

/**
 * Office-style formula bar: [ NameBox ] [ × ✓ fx ] [ formula input ]
 * Now with function autocomplete — type `=SU` to see SUM / SUMIF / etc.
 */
export function FormulaBar() {
  const api = useUniverAPI();
  const { ready, a1, displayValue } = useActiveCellState();

  const [draft, setDraft] = useState<string | null>(null);
  const editing = draft !== null;
  const inputRef = useRef<HTMLInputElement>(null);

  const [suggestions, setSuggestions] = useState<FormulaFn[]>([]);
  const [selectedIdx, setSelectedIdx] = useState(0);
  const [suggestionsAnchor, setSuggestionsAnchor] = useState<DOMRect | null>(null);

  const draftCellRef = useRef<string | null>(null);

  // Excel-style range picker: while editing a formula (draft starts with
  // `=`), clicking a sheet tab or a cell on the grid should insert a
  // reference at the caret instead of committing the draft. The origin
  // — the sheet + cell where editing started — is captured up front so
  // Enter commits back to it even after the user has been clicking
  // around other sheets.
  const isFormulaEdit = editing && (draft ?? '').startsWith('=');
  const originSheetIdRef = useRef<string | null>(null);
  const originRowRef = useRef<number | null>(null);
  const originColRef = useRef<number | null>(null);
  /** Set while we programmatically restore the origin sheet+cell on
   *  commit/cancel, so the SelectionChanged listener doesn't loop back
   *  and insert the origin's own ref into the draft. */
  const restoringOriginRef = useRef(false);
  /** Caret position the user last had in the input — read from the live
   *  input when it has focus, falls back to this ref while focus is
   *  elsewhere (canvas, sheet tab). Without this fallback the picker
   *  inserts every ref at position 0 once the input has blurred. */
  const lastCaretRef = useRef<number>(0);

  useEffect(() => {
    if (!editing) return;
    // In picker mode the active cell intentionally changes while the
    // user picks ranges — don't auto-commit on that change. Commit
    // happens only via Enter / Tab / Shift+Tab / commit button.
    if (isFormulaEdit) return;
    if (draftCellRef.current && draftCellRef.current !== a1 && api) {
      const original = draftCellRef.current;
      const text = draft ?? '';
      const wb = api.getActiveWorkbook();
      const sheet = wb?.getActiveSheet();
      if (wb && sheet) {
        const originalRange = sheet.getRange(original);
        const currentRange = sheet.getActiveRange();
        originalRange.activate();
        commitToActiveCell(api, text);
        currentRange?.activate();
      }
      setDraft(null);
      draftCellRef.current = null;
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [a1]);

  // Capture the origin sheet+cell the moment formula edit begins.
  useEffect(() => {
    if (!isFormulaEdit) {
      originSheetIdRef.current = null;
      originRowRef.current = null;
      originColRef.current = null;
      return;
    }
    if (originSheetIdRef.current !== null || !api) return;
    const wb = api.getActiveWorkbook();
    const sheet = wb?.getActiveSheet();
    const range = sheet?.getActiveRange();
    if (!sheet || !range) return;
    originSheetIdRef.current = sheet.getSheetId();
    originRowRef.current = range.getRow();
    originColRef.current = range.getColumn();
  }, [isFormulaEdit, api]);

  // Track the live caret position so the picker can splice refs in
  // even after the canvas has stolen focus. `selectionStart` returns
  // null on a blurred input in some browsers, so we cache the last
  // known value on every input event.
  const trackCaret = useCallback(() => {
    const input = inputRef.current;
    if (!input) return;
    const c = input.selectionStart;
    if (typeof c === 'number') lastCaretRef.current = c;
  }, []);

  // Listen for canvas-driven selection changes (sheet-ui's
  // `SelectionChanged` event fires only on user input — programmatic
  // `range.activate()` calls go through a different code path, so our
  // commit/cancel restore won't trigger this).
  useEffect(() => {
    if (!api || !isFormulaEdit) return;
    const disp = api.addEvent(api.Event.SelectionChanged, () => {
      if (restoringOriginRef.current) return;
      const wb = api.getActiveWorkbook();
      const sheet = wb?.getActiveSheet();
      const range = sheet?.getActiveRange();
      if (!wb || !sheet || !range) return;
      // Skip if the user is at the origin cell on the origin sheet —
      // that's just "click to refocus the start", not a pick.
      const sheetId = sheet.getSheetId();
      if (
        sheetId === originSheetIdRef.current &&
        range.getRow() === originRowRef.current &&
        range.getColumn() === originColRef.current &&
        range.getWidth() === 1 &&
        range.getHeight() === 1
      ) {
        return;
      }
      const caret = lastCaretRef.current;
      const currentDraft = draft ?? '';
      if (!canInsertRefAtCaret(currentDraft, caret)) return;
      const localA1 = range.getA1Notation();
      const refStr =
        sheetId === originSheetIdRef.current
          ? localA1
          : `${quoteSheetName(sheet.getSheetName())}!${localA1}`;
      const next = insertRefAtCaret(currentDraft, caret, refStr);
      flushSync(() => setDraft(next.value));
      lastCaretRef.current = next.caret;
      // Refocus the input so the user can keep typing operators / commit
      // with Enter. requestAnimationFrame so the focus call lands after
      // React's state flush settles.
      requestAnimationFrame(() => {
        inputRef.current?.focus();
        inputRef.current?.setSelectionRange(next.caret, next.caret);
      });
    });
    return () => disp.dispose();
  }, [api, isFormulaEdit, draft]);

  const value = editing ? (draft ?? '') : displayValue;

  // Recompute suggestions whenever the value or caret changes during editing.
  useLayoutEffect(() => {
    if (!editing || !inputRef.current) {
      setSuggestions([]);
      return;
    }
    const caret = inputRef.current.selectionStart ?? value.length;
    const frag = extractFunctionFragment(value, caret);
    const next = frag ? suggestFunctions(frag) : [];
    setSuggestions(next);
    setSelectedIdx(0);
    if (next.length > 0) {
      setSuggestionsAnchor(inputRef.current.getBoundingClientRect());
    }
  }, [value, editing]);

  const onChange = (e: ChangeEvent<HTMLInputElement>) => {
    if (!editing) draftCellRef.current = a1;
    setDraft(e.target.value);
  };

  const insertSuggestion = (fn: FormulaFn) => {
    if (!inputRef.current) return;
    const caret = inputRef.current.selectionStart ?? value.length;
    const frag = extractFunctionFragment(value, caret) ?? '';
    const before = value.slice(0, caret - frag.length);
    const after = value.slice(caret);
    const insertion = `${fn.name}(`;
    const next = `${before}${insertion}${after}`;
    const nextCaret = before.length + insertion.length;

    flushSync(() => setDraft(next));
    setSuggestions([]);
    requestAnimationFrame(() => {
      inputRef.current?.setSelectionRange(nextCaret, nextCaret);
      inputRef.current?.focus();
    });
  };

  /** Programmatically switch back to the origin sheet+cell that was
   *  active when the user started typing the formula. Used by commit
   *  and revert to undo any sheet-tab / cell clicks the picker made.
   *  Wraps the work in `restoringOriginRef` so the SelectionChanged
   *  listener treats this as a no-op instead of inserting the origin's
   *  own ref into the (already-finalized) draft. */
  const restoreOrigin = (): void => {
    if (!api || originSheetIdRef.current === null) return;
    const wb = api.getActiveWorkbook();
    if (!wb) return;
    const sheet = wb.getSheets().find((s) => s.getSheetId() === originSheetIdRef.current);
    if (!sheet) return;
    restoringOriginRef.current = true;
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (wb as any).setActiveSheet(sheet);
      const r = originRowRef.current ?? 0;
      const c = originColRef.current ?? 0;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (sheet as any).getRange(r, c).activate();
    } finally {
      // Release on the next tick — Univer's selection update may fire
      // async, so we keep the guard up for one frame.
      requestAnimationFrame(() => {
        restoringOriginRef.current = false;
      });
    }
  };

  const commit = (direction: CommitDirection = 'none') => {
    if (!api || !editing) return;
    const text = draft ?? '';
    restoreOrigin();
    commitToActiveCell(api, text, direction);
    flushSync(() => setDraft(null));
    setSuggestions([]);
    draftCellRef.current = null;
    originSheetIdRef.current = null;
    originRowRef.current = null;
    originColRef.current = null;
  };

  const revert = () => {
    restoreOrigin();
    flushSync(() => setDraft(null));
    setSuggestions([]);
    draftCellRef.current = null;
    originSheetIdRef.current = null;
    originRowRef.current = null;
    originColRef.current = null;
  };

  const onKeyDown = (e: KeyboardEvent<HTMLInputElement>) => {
    if (suggestions.length > 0) {
      if (e.key === 'ArrowDown') {
        e.preventDefault();
        setSelectedIdx((i) => (i + 1) % suggestions.length);
        return;
      }
      if (e.key === 'ArrowUp') {
        e.preventDefault();
        setSelectedIdx((i) => (i - 1 + suggestions.length) % suggestions.length);
        return;
      }
      // Tab / Enter with an open suggestion list ⇒ accept the selected
      // suggestion (Excel-equivalent). Shift+Tab and Shift+Enter still
      // commit-and-navigate — don't hijack those.
      if ((e.key === 'Tab' && !e.shiftKey) || (e.key === 'Enter' && !e.shiftKey)) {
        e.preventDefault();
        insertSuggestion(suggestions[selectedIdx]);
        return;
      }
      if (e.key === 'Escape') {
        e.preventDefault();
        setSuggestions([]);
        return;
      }
    }

    // F4 — cycle absolute/relative on the ref under the caret. Excel's
    // muscle memory. Works mid-edit (formula or plain text); a no-op
    // if there's no recognizable ref under the caret.
    if (e.key === 'F4' && editing) {
      e.preventDefault();
      const input = inputRef.current;
      if (!input) return;
      const caret = input.selectionStart ?? (draft ?? '').length;
      const rewrite = cycleAbsoluteRefAtCaret(draft ?? '', caret);
      if (!rewrite) return;
      flushSync(() => setDraft(rewrite.value));
      requestAnimationFrame(() => {
        input.setSelectionRange(rewrite.caret, rewrite.caret);
        input.focus();
      });
      return;
    }

    if (e.key === 'Enter') {
      e.preventDefault();
      commit(e.shiftKey ? 'up' : 'down');
      inputRef.current?.blur();
    } else if (e.key === 'Tab') {
      // Excel: Tab commits and moves right, Shift+Tab commits and moves
      // left. Without this handler the browser's focus trap would steal
      // Tab and leave the value uncommitted.
      e.preventDefault();
      commit(e.shiftKey ? 'left' : 'right');
      inputRef.current?.blur();
    } else if (e.key === 'Escape') {
      e.preventDefault();
      revert();
      inputRef.current?.blur();
    }
  };

  return (
    <div className="formula-bar" data-testid="formula-bar">
      <NameBox a1={a1} />


      <div className="formula-bar__actions" role="group" aria-label="Formula bar actions">
        <Tooltip label="Cancel (Esc)">
          <button
            type="button"
            className="btn btn--icon formula-bar__action"
            data-testid="formula-cancel"
            aria-label="Cancel"
            disabled={!editing}
            onMouseDown={(e) => e.preventDefault()}
            onClick={revert}
          >
            <Icon name="close" size="sm" />
          </button>
        </Tooltip>
        <Tooltip label="Enter (↵)">
          <button
            type="button"
            className="btn btn--icon formula-bar__action"
            data-testid="formula-commit"
            aria-label="Enter"
            disabled={!editing}
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => {
              commit();
              inputRef.current?.blur();
            }}
          >
            <Icon name="check" size="sm" />
          </button>
        </Tooltip>
        <Tooltip label="Insert function">
          <span
            className="formula-bar__fx"
            data-testid="formula-fx"
            aria-label="Insert function"
          >
            fx
          </span>
        </Tooltip>
      </div>

      <input
        ref={inputRef}
        type="text"
        className="formula-bar__input"
        data-testid="formula-input"
        aria-label="Formula input"
        spellCheck={false}
        value={value}
        disabled={!ready}
        onChange={(e) => {
          trackCaret();
          onChange(e);
        }}
        onKeyDown={onKeyDown}
        onKeyUp={trackCaret}
        onClick={trackCaret}
        onSelect={() => {
          if (!editing || !inputRef.current) return;
          const caret = inputRef.current.selectionStart ?? value.length;
          lastCaretRef.current = caret;
          // Recompute suggestions when the user moves the caret with arrows.
          const frag = extractFunctionFragment(value, caret);
          setSuggestions(frag ? suggestFunctions(frag) : []);
        }}
        onBlur={(e) => {
          // Don't dismiss when the user clicks a suggestion (focus moves
          // briefly into the popover). Re-focus heuristic: check if the
          // relatedTarget is the suggestion list.
          if ((e.relatedTarget as HTMLElement | null)?.closest('[data-testid="formula-suggestions"]')) return;
          // In picker mode (formula edit) the user is clicking the
          // grid / sheet tabs to build refs — blur is expected. The
          // SelectionChanged listener splices refs in; commit happens
          // only on Enter/Tab/Esc, not on focus loss.
          if (isFormulaEdit) return;
          if (editing) commit();
        }}
      />

      {editing && suggestions.length > 0 && suggestionsAnchor && (
        <ul
          className="formula-suggestions"
          data-testid="formula-suggestions"
          role="listbox"
          style={{
            top: suggestionsAnchor.bottom + 2,
            left: suggestionsAnchor.left,
            width: Math.max(280, suggestionsAnchor.width / 2),
          }}
        >
          {suggestions.map((fn, i) => (
            <li
              key={fn.name}
              role="option"
              aria-selected={i === selectedIdx}
              className={`formula-suggestions__item${i === selectedIdx ? ' formula-suggestions__item--selected' : ''}`}
              data-testid={`formula-suggestion-${fn.name}`}
              onMouseDown={(e) => {
                // mousedown not click — so the input doesn't blur first.
                e.preventDefault();
                insertSuggestion(fn);
              }}
            >
              <span className="formula-suggestions__name">{fn.name}</span>
              <span className="formula-suggestions__desc">{fn.description}</span>
            </li>
          ))}
        </ul>
      )}
    </div>
  );
}

/**
 * Editable A1-reference input. Defaults to the active cell address. Accepts
 * any A1 notation the Univer facade's `getRange()` understands — single cell
 * (`B5`), range (`B5:D10`), or column/row (`B:B`). On Enter, parses and
 * activates the range; on Esc, reverts to the current address. Invalid
 * references snap back without changing the selection.
 */
function NameBox({ a1 }: { a1: string }) {
  const api = useUniverAPI();
  const [draft, setDraft] = useState<string | null>(null);
  const inputRef = useRef<HTMLInputElement>(null);
  const value = draft ?? a1;

  const commit = () => {
    const target = (draft ?? '').trim();
    setDraft(null);
    if (!api || !target || target === a1) return;
    const wb = api.getActiveWorkbook();
    const sheet = wb?.getActiveSheet();
    if (!sheet) return;
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const range = (sheet as any).getRange(target);
      range?.activate?.();
    } catch {
      /* invalid reference — silent, the input has already snapped back */
    }
  };

  return (
    <input
      ref={inputRef}
      type="text"
      className="formula-bar__name-box"
      data-testid="name-box"
      title="Name Box — type a cell reference (e.g. B5 or B5:D10) and press Enter"
      aria-label="Name Box"
      spellCheck={false}
      value={value || ''}
      onChange={(e) => setDraft(e.target.value)}
      onFocus={(e) => {
        setDraft(a1);
        // Select the existing reference so typing replaces it immediately —
        // matches Excel's "click name box, type, Enter" muscle memory.
        requestAnimationFrame(() => e.target.select());
      }}
      onBlur={commit}
      onKeyDown={(e) => {
        if (e.key === 'Enter') {
          e.preventDefault();
          commit();
          inputRef.current?.blur();
        } else if (e.key === 'Escape') {
          e.preventDefault();
          setDraft(null);
          inputRef.current?.blur();
        }
      }}
    />
  );
}
