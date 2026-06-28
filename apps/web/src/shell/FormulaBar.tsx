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
import { Dialog } from './Dialog';
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
  expandFunctionArgsAtCaret,
  extractFunctionFragment,
  FORMULA_FUNCTIONS,
  getFunctionSignature,
  suggestFunctions,
  suggestSheetNames,
  type FormulaFn,
} from './formula-functions';

/**
 * Suggestion list entry — a function call (insert `NAME(`) or a sheet
 * name (insert `Name!`). Keyboard nav + click handling treats them
 * uniformly; only the display + insertion text differ.
 */
type Suggestion =
  | { kind: 'fn'; fn: FormulaFn }
  | { kind: 'sheet'; name: string };

/**
 * Office-style formula bar: [ NameBox ] [ × ✓ fx ] [ formula input ]
 * Now with function autocomplete — type `=SU` to see SUM / SUMIF / etc.
 */
export function FormulaBar() {
  const api = useUniverAPI();
  const { ready, a1, displayValue, isMultiCell, selRows, selCols } = useActiveCellState();

  const [draft, setDraft] = useState<string | null>(null);
  const editing = draft !== null;
  const inputRef = useRef<HTMLInputElement>(null);

  const [suggestions, setSuggestions] = useState<Suggestion[]>([]);
  const [selectedIdx, setSelectedIdx] = useState(0);
  const [suggestionsAnchor, setSuggestionsAnchor] = useState<DOMRect | null>(null);
  const [showInsertFunction, setShowInsertFunction] = useState(false);
  const [functionSearch, setFunctionSearch] = useState('');
  const [addToSelectionMode, setAddToSelectionMode] = useState(false);

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
  /** Bounds of the last reference the picker spliced in. While
   *  non-null, the NEXT picker fire REPLACES this region instead of
   *  inserting after it — that's the Excel behaviour where dragging
   *  a range updates the formula in-place rather than concatenating
   *  every cursor stop. Cleared when the user types in the formula
   *  bar (any character finalises the current ref). */
  const lastPickRangeRef = useRef<{ start: number; end: number } | null>(null);
  /** True for the one onChange that follows a picker-driven setDraft —
   *  prevents the input's onChange from clearing `lastPickRangeRef`
   *  for the picker's own write. */
  const pickerWroteRef = useRef(false);

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
      lastPickRangeRef.current = null;
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
    // Subscribe to BOTH SelectionChanged (Univer's programmatic /
    // internal selection updates) and SelectionMoveEnd (the event
    // that fires reliably on user drag-release on the canvas).
    // SelectionChanged alone leaves real user clicks silent — a
    // demo user reported the cross-sheet picker doing nothing
    // before we wired SelectionMoveEnd here too.
    const onPick = () => {
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
      const insertCaret = lastPickRangeRef.current?.start ?? caret;
      if (!canInsertRefAtCaret(currentDraft, insertCaret)) return;
      const localA1 = range.getA1Notation();
      const refStr =
        sheetId === originSheetIdRef.current
          ? localA1
          : `${quoteSheetName(sheet.getSheetName())}!${localA1}`;

      // If we already inserted a ref this picker cycle, REPLACE that
      // region instead of appending — matches Excel's drag-to-update-
      // ref behaviour.
      let nextValue: string;
      let nextCaret: number;
      if (lastPickRangeRef.current) {
        const { start, end } = lastPickRangeRef.current;
        nextValue = currentDraft.slice(0, start) + refStr + currentDraft.slice(end);
        nextCaret = start + refStr.length;
      } else {
        const spliced = insertRefAtCaret(currentDraft, insertCaret, refStr);
        nextValue = spliced.value;
        nextCaret = spliced.caret;
      }
      lastPickRangeRef.current = { start: nextCaret - refStr.length, end: nextCaret };
      pickerWroteRef.current = true;
      lastCaretRef.current = nextCaret;
      // Defer the state flush to a microtask — calling flushSync from
      // inside the Univer event listener races React's render phase
      // and triggers the "flushSync called from a lifecycle method"
      // warning. queueMicrotask runs after the current render settles.
      queueMicrotask(() => {
        setDraft(nextValue);
        requestAnimationFrame(() => {
          inputRef.current?.focus();
          inputRef.current?.setSelectionRange(nextCaret, nextCaret);
        });
      });
    };
    const d1 = api.addEvent(api.Event.SelectionChanged, onPick);
    const d2 = api.addEvent(api.Event.SelectionMoveEnd, onPick);
    return () => {
      d1.dispose();
      d2.dispose();
    };
  }, [api, isFormulaEdit, draft]);

  const value = editing ? (draft ?? '') : displayValue;

  // Recompute suggestions whenever the value or caret changes during
  // editing. Mix function names + sheet names — both are valid things
  // to type after `=` (functions go first, sheet names below them with
  // a `Sheet` badge so users can tell them apart).
  useLayoutEffect(() => {
    if (!editing || !inputRef.current) {
      setSuggestions([]);
      return;
    }
    const caret = inputRef.current.selectionStart ?? value.length;
    const frag = extractFunctionFragment(value, caret);
    if (!frag) {
      setSuggestions([]);
      return;
    }
    const fnList = suggestFunctions(frag);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const sheets = ((api?.getActiveWorkbook?.()?.getSheets?.() as any[]) ?? [])
      .map((s) => s.getSheetName?.())
      .filter((n): n is string => typeof n === 'string');
    const sheetList = suggestSheetNames(frag, sheets);
    const next: Suggestion[] = [
      ...fnList.map((fn): Suggestion => ({ kind: 'fn', fn })),
      ...sheetList.map((name): Suggestion => ({ kind: 'sheet', name })),
    ];
    setSuggestions(next);
    setSelectedIdx(0);
    if (next.length > 0) {
      setSuggestionsAnchor(inputRef.current.getBoundingClientRect());
    }
  }, [value, editing, api]);

  const onChange = (e: ChangeEvent<HTMLInputElement>) => {
    if (!editing) draftCellRef.current = a1;
    // Any user-driven character finalises the in-progress picker ref —
    // typing `,` after `=SUM(A1:A3` means the next pick starts a new
    // ref instead of replacing the old one. Skip when this onChange
    // is the echo of our own picker write.
    if (!pickerWroteRef.current) {
      lastPickRangeRef.current = null;
    }
    pickerWroteRef.current = false;
    setDraft(e.target.value);
  };

  const startFormulaEdit = useCallback(() => {
    if (!editing) {
      draftCellRef.current = a1;
      flushSync(() => setDraft('='));
      requestAnimationFrame(() => {
        inputRef.current?.focus();
        inputRef.current?.setSelectionRange(1, 1);
        lastCaretRef.current = 1;
      });
      return;
    }
    requestAnimationFrame(() => {
      const input = inputRef.current;
      if (!input) return;
      input.focus();
      const caret = input.selectionStart ?? (draft ?? '').length;
      input.setSelectionRange(caret, caret);
      lastCaretRef.current = caret;
    });
  }, [editing, a1, draft]);

  const insertSuggestion = (s: Suggestion) => {
    if (!inputRef.current) return;
    const caret = inputRef.current.selectionStart ?? value.length;
    const frag = extractFunctionFragment(value, caret) ?? '';
    const before = value.slice(0, caret - frag.length);
    const after = value.slice(caret);
    // Functions insert `NAME(` (open paren for argument typing).
    // Sheet names insert `Name!` (the bang for the cell ref that
    // follows).
    const insertion = s.kind === 'fn' ? `${s.fn.name}(` : `${s.name}!`;
    const next = `${before}${insertion}${after}`;
    const nextCaret = before.length + insertion.length;

    flushSync(() => setDraft(next));
    setSuggestions([]);
    requestAnimationFrame(() => {
      inputRef.current?.setSelectionRange(nextCaret, nextCaret);
      inputRef.current?.focus();
    });
  };

  const insertFunctionFromDialog = useCallback(
    (fn: FormulaFn) => {
      const input = inputRef.current;
      const current = editing ? (draft ?? '') : '=';
      const base = current.startsWith('=') ? current : `=${current}`;
      const caret = input?.selectionStart ?? base.length;
      const frag = extractFunctionFragment(base, caret) ?? '';
      const before = base.slice(0, caret - frag.length);
      const after = base.slice(caret);
      const insertion = `${fn.name}()`;
      const next = `${before}${insertion}${after}`;
      const nextCaret = before.length + fn.name.length + 1;
      draftCellRef.current = a1;
      flushSync(() => {
        setDraft(next);
        setShowInsertFunction(false);
      });
      requestAnimationFrame(() => {
        inputRef.current?.focus();
        inputRef.current?.setSelectionRange(nextCaret, nextCaret);
        lastCaretRef.current = nextCaret;
      });
    },
    [editing, draft, a1],
  );

  const insertFunctionArgs = useCallback(() => {
    const input = inputRef.current;
    if (!input || !editing) return;
    const caret = input.selectionStart ?? (draft ?? '').length;
    const expanded = expandFunctionArgsAtCaret(draft ?? '', caret);
    if (!expanded) return;
    flushSync(() => setDraft(expanded.value));
    requestAnimationFrame(() => {
      inputRef.current?.focus();
      inputRef.current?.setSelectionRange(expanded.selectionStart, expanded.selectionEnd);
      lastCaretRef.current = expanded.selectionEnd;
    });
  }, [editing, draft]);

  useEffect(() => {
    const openInsertFunction = () => {
      startFormulaEdit();
      setShowInsertFunction(true);
      setFunctionSearch('');
    };
    const openFunctionArgs = () => {
      if (!editing) startFormulaEdit();
      requestAnimationFrame(() => insertFunctionArgs());
    };
    document.addEventListener('casual-open-insert-function', openInsertFunction);
    document.addEventListener('casual-insert-function-args', openFunctionArgs);
    return () => {
      document.removeEventListener('casual-open-insert-function', openInsertFunction);
      document.removeEventListener('casual-insert-function-args', openFunctionArgs);
    };
  }, [startFormulaEdit, insertFunctionArgs, editing]);

  useEffect(() => {
    const onModeChange = (event: Event) => {
      const active = (event as CustomEvent<{ active?: boolean }>).detail?.active;
      setAddToSelectionMode(Boolean(active));
    };
    document.addEventListener('casual-add-to-selection-mode-changed', onModeChange);
    return () => document.removeEventListener('casual-add-to-selection-mode-changed', onModeChange);
  }, []);

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
    if (e.shiftKey && !e.altKey && e.key === 'F3') {
      e.preventDefault();
      setShowInsertFunction(true);
      setFunctionSearch('');
      return;
    }

    if ((e.metaKey || e.ctrlKey) && e.shiftKey && !e.altKey && e.key.toLowerCase() === 'a') {
      e.preventDefault();
      insertFunctionArgs();
      return;
    }

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
      // Univer 0.25 binds F4 → RepeatLastActionCommand at document level with a
      // `whenSheetEditorFocused` precondition, which is true while the formula bar
      // drives the background cell editor. Stop the native event so only our
      // ref-under-caret cycling runs (Excel parity) — React's synthetic
      // stopPropagation alone won't reach Univer's native listener.
      e.nativeEvent.stopImmediatePropagation();
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
      {addToSelectionMode && (
        <span
          className="formula-bar__selection-mode"
          data-testid="add-to-selection-indicator"
          title="Add non-adjacent ranges to the current selection"
        >
          Add to Selection
        </span>
      )}
      {isMultiCell && (
        <span
          className="formula-bar__sel-dims"
          data-testid="sel-dimensions"
          title="Selection dimensions"
        >
          {selRows}R × {selCols}C
        </span>
      )}


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
          <button
            type="button"
            className="formula-bar__fx"
            data-testid="formula-fx"
            aria-label="Insert function"
            onClick={() => {
              startFormulaEdit();
              setShowInsertFunction(true);
              setFunctionSearch('');
            }}
          >
            fx
          </button>
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
        onPaste={(e) => {
          // Multi-line formulas are valid for readability — Excel
          // allows newlines inside the formula bar for SUM/IF style
          // wrapping. Native <input type="text"> drops newlines in
          // some browsers and replaces them with U+000A in others;
          // either way the visible single-line input can't render
          // them. Normalize newlines to spaces so the *formula* stays
          // valid even though the editor doesn't show line breaks.
          //
          // (Switching to <textarea> would let the bar show real
          // line breaks, but that breaks the existing Enter-commits
          // / Tab-commits semantics across the suite — bigger change
          // than this bug warrants.)
          const text = e.clipboardData?.getData('text/plain');
          if (!text || !/[\r\n]/.test(text)) return;
          e.preventDefault();
          const normalized = text.replace(/\r\n?|\n/g, ' ');
          const input = inputRef.current;
          if (!input) return;
          const start = input.selectionStart ?? value.length;
          const end = input.selectionEnd ?? start;
          const before = value.slice(0, start);
          const after = value.slice(end);
          const next = `${before}${normalized}${after}`;
          if (!editing) draftCellRef.current = a1;
          flushSync(() => setDraft(next));
          requestAnimationFrame(() => {
            const caret = before.length + normalized.length;
            input.setSelectionRange(caret, caret);
            lastCaretRef.current = caret;
          });
        }}
        onSelect={() => {
          if (!editing || !inputRef.current) return;
          const caret = inputRef.current.selectionStart ?? value.length;
          lastCaretRef.current = caret;
          // Recompute suggestions when the user moves the caret with arrows.
          // Mirrors the main recompute effect — keep both in lockstep.
          const frag = extractFunctionFragment(value, caret);
          if (!frag) {
            setSuggestions([]);
            return;
          }
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const sheets = ((api?.getActiveWorkbook?.()?.getSheets?.() as any[]) ?? [])
            .map((s) => s.getSheetName?.())
            .filter((n): n is string => typeof n === 'string');
          setSuggestions([
            ...suggestFunctions(frag).map((fn): Suggestion => ({ kind: 'fn', fn })),
            ...suggestSheetNames(frag, sheets).map((name): Suggestion => ({ kind: 'sheet', name })),
          ]);
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
          {suggestions.map((s, i) => {
            const name = s.kind === 'fn' ? s.fn.name : s.name;
            const desc = s.kind === 'fn' ? s.fn.description : 'Sheet';
            return (
              <li
                key={`${s.kind}-${name}`}
                role="option"
                aria-selected={i === selectedIdx}
                className={`formula-suggestions__item${i === selectedIdx ? ' formula-suggestions__item--selected' : ''}${s.kind === 'sheet' ? ' formula-suggestions__item--sheet' : ''}`}
                data-testid={`formula-suggestion-${name}`}
                data-kind={s.kind}
                onMouseDown={(e) => {
                  // mousedown not click — so the input doesn't blur first.
                  e.preventDefault();
                  insertSuggestion(s);
                }}
              >
                <span className="formula-suggestions__name">{name}</span>
                <span className="formula-suggestions__desc">{desc}</span>
              </li>
            );
          })}
        </ul>
      )}

      {showInsertFunction && (
        <InsertFunctionDialog
          search={functionSearch}
          onSearch={setFunctionSearch}
          onClose={() => setShowInsertFunction(false)}
          onChoose={insertFunctionFromDialog}
        />
      )}
    </div>
  );
}

function InsertFunctionDialog({
  search,
  onSearch,
  onClose,
  onChoose,
}: {
  search: string;
  onSearch: (value: string) => void;
  onClose: () => void;
  onChoose: (fn: FormulaFn) => void;
}) {
  const matches = FORMULA_FUNCTIONS.filter((fn) => {
    const q = search.trim().toLowerCase();
    if (!q) return true;
    return fn.name.toLowerCase().includes(q) || fn.description.toLowerCase().includes(q);
  }).slice(0, 40);
  const [selected, setSelected] = useState(0);

  useEffect(() => {
    setSelected(0);
  }, [search]);

  const choice = matches[selected] ?? matches[0] ?? null;

  return (
    <Dialog
      title="Insert Function"
      onClose={onClose}
      data-testid="insert-function-dialog"
      footer={
        <>
          <button type="button" className="btn-secondary" onClick={onClose}>
            Cancel
          </button>
          <button
            type="button"
            className="btn-primary"
            data-testid="insert-function-apply"
            disabled={!choice}
            onClick={() => choice && onChoose(choice)}
          >
            Insert
          </button>
        </>
      }
    >
      <div className="insert-function">
        <input
          className="input"
          data-testid="insert-function-search"
          aria-label="Search functions"
          placeholder="Search functions"
          value={search}
          onChange={(e) => onSearch(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === 'ArrowDown') {
              e.preventDefault();
              setSelected((i) => Math.min(matches.length - 1, i + 1));
            } else if (e.key === 'ArrowUp') {
              e.preventDefault();
              setSelected((i) => Math.max(0, i - 1));
            } else if (e.key === 'Enter') {
              e.preventDefault();
              if (choice) onChoose(choice);
            }
          }}
        />
        <div className="insert-function__body">
          <ul className="insert-function__list" data-testid="insert-function-list" role="listbox">
            {matches.map((fn, i) => (
              <li
                key={fn.name}
                role="option"
                aria-selected={i === selected}
                className={`insert-function__item${i === selected ? ' insert-function__item--active' : ''}`}
                data-testid={`insert-function-item-${fn.name}`}
                onMouseDown={(e) => {
                  e.preventDefault();
                  onChoose(fn);
                }}
                onMouseEnter={() => setSelected(i)}
              >
                <div className="insert-function__name">{fn.name}</div>
                <div className="insert-function__desc">{fn.description}</div>
              </li>
            ))}
          </ul>
          <div className="insert-function__detail" data-testid="insert-function-detail">
            {choice ? (
              <>
                <div className="insert-function__detail-name">{choice.name}</div>
                <div className="insert-function__detail-sig">{getFunctionSignature(choice)}</div>
                <p className="insert-function__detail-desc">{choice.description}</p>
              </>
            ) : (
              <p className="insert-function__detail-desc">No matching functions.</p>
            )}
          </div>
        </div>
      </div>
    </Dialog>
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

  useEffect(() => {
    const focusNameBox = () => {
      const input = inputRef.current;
      if (!input) return;
      setDraft(a1);
      requestAnimationFrame(() => {
        input.focus();
        input.select();
      });
    };
    document.addEventListener('casual-focus-name-box', focusNameBox);
    return () => document.removeEventListener('casual-focus-name-box', focusNameBox);
  }, [a1]);

  const commit = () => {
    const target = (draft ?? '').trim();
    setDraft(null);
    if (!api || !target || target === a1) return;
    const wb = api.getActiveWorkbook();
    const sheet = wb?.getActiveSheet();
    if (!sheet) return;
    if (document.body.dataset.addToSelectionMode === 'true') {
      document.dispatchEvent(
        new CustomEvent('casual-add-selection-a1', {
          detail: { target },
        }),
      );
      return;
    }
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
