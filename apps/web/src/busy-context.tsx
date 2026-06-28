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

import { createContext, useCallback, useContext, useMemo, useRef, useState } from 'react';
import type { ReactNode } from 'react';

/**
 * Lightweight "the app is doing something" indicator. Drives the small
 * pill in the title bar so a user clicking Format as Table on a 100k
 * row sheet doesn't think the app froze — until we modify Univer's
 * mutation pipeline for actual perf, this is the honest UX patch.
 *
 * Imperative on purpose: action handlers call `runBusy('Creating
 * table…', () => doTheSyncWork())` and the helper:
 *   1. sets the label
 *   2. yields a frame so the pill paints BEFORE the sync work blocks
 *   3. runs the work
 *   4. clears the label
 *
 * No automatic main-thread-blocked detection — JS can't observe its
 * own freeze. We only show the pill when an action explicitly asks
 * us to.
 */
type BusyState = { label: string } | null;

export type BusyCtxValue = {
  state: BusyState;
  setBusy: (label: string | null) => void;
  /**
   * Helper that wraps a slow-but-synchronous operation in the
   * setBusy → yield-a-frame → run → clearBusy dance. Always async,
   * always clears the label on exit (even when `fn` throws), so the
   * pill can't get stuck visible after a runtime error.
   */
  runBusy: <T>(label: string, fn: () => T | Promise<T>) => Promise<T>;
};

export const BusyContext = createContext<BusyCtxValue>({
  state: null,
  setBusy: () => undefined,
  runBusy: async (_label, fn) => fn(),
});

export function useBusy(): BusyCtxValue {
  return useContext(BusyContext);
}

export function BusyProvider({ children }: { children: ReactNode }) {
  const [state, setState] = useState<BusyState>(null);
  // Latest-label-wins: keep a ref so rapid setBusy calls during a
  // burst of nested actions don't drop the most recent label.
  const labelRef = useRef<string | null>(null);

  const setBusy = useCallback((label: string | null) => {
    labelRef.current = label;
    setState(label ? { label } : null);
  }, []);

  // Dev-only window hook so e2e can drive the pill deterministically
  // without racing the fast-resolving work of small Format-as-Table
  // fixtures. Tree-shaken in production via import.meta.env.DEV.
  if (import.meta.env.DEV) {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (globalThis as any).__setBusy = setBusy;
  }

  const runBusy = useCallback(<T,>(label: string, fn: () => T | Promise<T>): Promise<T> => {
    return new Promise<T>((resolve, reject) => {
      labelRef.current = label;
      setState({ label });
      // Two animation frames: the first commits the React render that
      // mounts the pill; the second yields the paint pass so the user
      // sees it before our synchronous work blocks the thread.
      requestAnimationFrame(() => {
        requestAnimationFrame(() => {
          try {
            const result = fn();
            Promise.resolve(result)
              .then((v) => {
                resolve(v);
              })
              .catch(reject)
              .finally(() => {
                // Only clear if no newer action took over (label
                // changed mid-flight = another runBusy in progress).
                if (labelRef.current === label) {
                  labelRef.current = null;
                  setState(null);
                }
              });
          } catch (err) {
            if (labelRef.current === label) {
              labelRef.current = null;
              setState(null);
            }
            reject(err);
          }
        });
      });
    });
  }, []);

  const value = useMemo<BusyCtxValue>(
    () => ({ state, setBusy, runBusy }),
    [state, setBusy, runBusy],
  );

  return <BusyContext.Provider value={value}>{children}</BusyContext.Provider>;
}
