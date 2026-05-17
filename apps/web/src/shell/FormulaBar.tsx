import {
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
import { commitToActiveCell } from './cell-actions';
import { Icon } from './Icon';
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

  useEffect(() => {
    if (!editing) return;
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

  const commit = () => {
    if (!api || !editing) return;
    commitToActiveCell(api, draft ?? '');
    flushSync(() => setDraft(null));
    setSuggestions([]);
    draftCellRef.current = null;
  };

  const revert = () => {
    flushSync(() => setDraft(null));
    setSuggestions([]);
    draftCellRef.current = null;
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
      if (e.key === 'Tab' || (e.key === 'Enter' && !e.shiftKey)) {
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

    if (e.key === 'Enter') {
      e.preventDefault();
      commit();
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
        <button
          type="button"
          className="btn btn--icon formula-bar__action"
          data-testid="formula-cancel"
          aria-label="Cancel"
          title="Cancel (Esc)"
          disabled={!editing}
          onMouseDown={(e) => e.preventDefault()}
          onClick={revert}
        >
          <Icon name="close" size="sm" />
        </button>
        <button
          type="button"
          className="btn btn--icon formula-bar__action"
          data-testid="formula-commit"
          aria-label="Enter"
          title="Enter (↵)"
          disabled={!editing}
          onMouseDown={(e) => e.preventDefault()}
          onClick={() => {
            commit();
            inputRef.current?.blur();
          }}
        >
          <Icon name="check" size="sm" />
        </button>
        <span
          className="formula-bar__fx"
          data-testid="formula-fx"
          aria-label="Insert function"
          title="Insert function"
        >
          fx
        </span>
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
        onChange={onChange}
        onKeyDown={onKeyDown}
        onSelect={() => {
          // Recompute suggestions when the user moves the caret with arrows.
          if (!editing || !inputRef.current) return;
          const caret = inputRef.current.selectionStart ?? value.length;
          const frag = extractFunctionFragment(value, caret);
          setSuggestions(frag ? suggestFunctions(frag) : []);
        }}
        onBlur={(e) => {
          // Don't dismiss when the user clicks a suggestion (focus moves
          // briefly into the popover). Re-focus heuristic: check if the
          // relatedTarget is the suggestion list.
          if ((e.relatedTarget as HTMLElement | null)?.closest('[data-testid="formula-suggestions"]')) return;
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
