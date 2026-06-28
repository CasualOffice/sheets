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
  createContext,
  useCallback,
  useContext,
  useMemo,
  useRef,
  useState,
  type ReactNode,
} from 'react';

/**
 * Save-status state. UX_AUDIT.md §4.3 / Phase 4 #16. Mirrors the
 * Google-Docs / Word-online "Saved 2 min ago" pill — the audit
 * called this out as a major missing confirmation surface.
 *
 * Lifecycle:
 *
 *   idle ──(start of save)──▶ saving ──(success)──▶ saved
 *                                    ╲              │
 *                                     ╲             │ (user edits — markUserEdited)
 *                                      ╲            ▼
 *                                       ─(failure)▶ error (with message + retry hint)
 *
 * The `saved` state carries `savedAt` so the pill can render
 * relative time ("just now" / "2 min ago" / "Saved at 14:32"). The
 * pill itself owns the tick; the context is pure data.
 */
export type SaveStatus =
  | { kind: 'idle' }
  | { kind: 'saving' }
  | { kind: 'saved'; savedAt: number }
  | { kind: 'error'; message: string };

export interface SaveStatusCtx {
  status: SaveStatus;
  /** Flip to `saving` — clears any previous error. Called at the top of `handleSave`. */
  markSaving: () => void;
  /** Flip to `saved` with timestamp `Date.now()`. Called at success tail. */
  markSaved: () => void;
  /** Flip to `error`. Called from a save's catch block. */
  markError: (message: string) => void;
  /**
   * Hint that the user has modified the workbook since the last save.
   * Resets `saved`/`error` to `idle` so the pill doesn't continue to
   * lie about "Saved 1 hour ago" when the user is mid-typing. No-op
   * during `saving` (the save itself races edits — settle on outcome
   * first).
   */
  markDirty: () => void;
}

const FALLBACK_CTX: SaveStatusCtx = {
  status: { kind: 'idle' },
  markSaving: () => undefined,
  markSaved: () => undefined,
  markError: () => undefined,
  markDirty: () => undefined,
};

const SaveStatusContext = createContext<SaveStatusCtx | null>(null);

export function SaveStatusProvider({ children }: { children: ReactNode }) {
  const [status, setStatus] = useState<SaveStatus>({ kind: 'idle' });
  // Keep the latest status in a ref so the callbacks don't get rebound
  // on every transition — TitleBar / EditTracker subscribe by reference.
  const ref = useRef(status);
  ref.current = status;

  const markSaving = useCallback(() => setStatus({ kind: 'saving' }), []);
  const markSaved = useCallback(
    () => setStatus({ kind: 'saved', savedAt: Date.now() }),
    [],
  );
  const markError = useCallback(
    (message: string) => setStatus({ kind: 'error', message }),
    [],
  );
  const markDirty = useCallback(() => {
    // Only knock the pill back to idle when there was a settled
    // result to invalidate. During `saving` the save will produce a
    // fresh terminal state on its own.
    if (ref.current.kind === 'saved' || ref.current.kind === 'error') {
      setStatus({ kind: 'idle' });
    }
  }, []);

  const value = useMemo<SaveStatusCtx>(
    () => ({ status, markSaving, markSaved, markError, markDirty }),
    [status, markSaving, markSaved, markError, markDirty],
  );

  return (
    <SaveStatusContext.Provider value={value}>{children}</SaveStatusContext.Provider>
  );
}

/**
 * Consumer hook. Returns a non-null fallback context when used outside
 * the provider — keeps WOPI / anonymous-room call sites that don't
 * mount the provider from crashing.
 */
export function useSaveStatus(): SaveStatusCtx {
  return useContext(SaveStatusContext) ?? FALLBACK_CTX;
}
