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
  useEffect,
  useMemo,
  useRef,
  useState,
} from 'react';
import type { ReactNode } from 'react';
import type { CasualSheetsAPI } from '../sheets/api';
import { readChartsFromSnapshot, writeChartsIntoSnapshot } from './resources';
import type { ChartModel } from './types';

/**
 * Chart store mirrored into `IWorkbookData.resources['__casual_sheets_charts__']`
 * at save time and re-hydrated when the active workbook changes. The store
 * keeps charts keyed by id (`ChartModel.id`); removal by id, updates by
 * replace — every change produces a new model object so React effect deps
 * work for downstream consumers (ChartLayer, ChartOverlay).
 *
 * SDK port: the app read/wrote the store through the hidden xlsx pre-pass +
 * `useWorkbook()` revision tracking. Here the store is loaded on mount and
 * persisted on every local change through the SDK's `api.getContent()` /
 * `api.setContent()` (which serialize/replace the active `IWorkbookData`).
 * `api` is also re-exported via `useCharts()` so downstream chart components
 * reach the FUniver facade through `api.univer` instead of a React hook.
 */
type ChartsCtxValue = {
  /** The SDK editor handle — chart components reach the FUniver facade via
   *  `api.univer` (never a React `useUniverAPI` hook, which doesn't exist in
   *  the SDK). */
  api: CasualSheetsAPI;
  charts: ChartModel[];
  /** Id of the chart with the selection frame + handles drawn. At most
   *  one chart is selected at a time (Excel single-select; multi-select
   *  via Ctrl+click is not implemented yet). */
  selectedId: string | null;
  insert: (chart: ChartModel) => void;
  remove: (id: string) => void;
  update: (id: string, patch: Partial<ChartModel>) => void;
  select: (id: string | null) => void;
  /** Replace the entire chart list. Used by CollabDriver to apply
   *  remote chart-state updates from the Yjs sync map. App code should
   *  prefer `insert / remove / update` — `__replaceAll` bypasses the
   *  collab broadcast tag, so consecutive calls from outside the
   *  collab driver can clobber each other. */
  __replaceAll: (next: ChartModel[], opts?: { fromCollab?: boolean }) => void;
  /** Subscribe to LOCAL chart-list changes (insert / remove / update,
   *  excluding remote echoes from `__replaceAll({ fromCollab: true })`).
   *  CollabDriver uses this to push our edits into the Yjs map without
   *  echoing them back. Returns an unsubscribe. */
  __subscribeLocal: (cb: (charts: ChartModel[]) => void) => () => void;
};

export const ChartsContext = createContext<ChartsCtxValue | null>(null);

export function useCharts(): ChartsCtxValue {
  const ctx = useContext(ChartsContext);
  if (!ctx) throw new Error('useCharts must be used inside <ChartsProvider>');
  return ctx;
}

export function ChartsProvider({ api, children }: { api: CasualSheetsAPI; children: ReactNode }) {
  const [charts, setCharts] = useState<ChartModel[]>(() =>
    readChartsFromSnapshot(api.getContent() ?? undefined),
  );
  const [selectedId, setSelectedId] = useState<string | null>(null);
  // Local change subscribers (used by CollabDriver to push to Yjs).
  // Stored on a ref so the subscribe API is stable across renders.
  const subsRef = useRef<Set<(c: ChartModel[]) => void>>(new Set());
  const notifyLocal = useCallback((next: ChartModel[]) => {
    for (const cb of subsRef.current) cb(next);
  }, []);

  // Persist the chart list into the active workbook snapshot on every local
  // change. Mirrors the app's write-into-resources pre-pass, but driven
  // through the SDK content accessors. Guarded by `persistingRef` so the
  // `setContent` round-trip we trigger here doesn't bounce back through the
  // `change` re-hydrate below as a phantom remote update.
  const persistingRef = useRef(false);
  const persist = useCallback(
    (next: ChartModel[]) => {
      const snap = api.getContent();
      if (!snap) return;
      writeChartsIntoSnapshot(snap, next);
      persistingRef.current = true;
      try {
        api.setContent(snap);
      } finally {
        persistingRef.current = false;
      }
    },
    [api],
  );

  // Re-hydrate on external content swaps (Open / New / collab remote-snapshot).
  // The app keyed this off `useWorkbook().meta.revision`; the SDK surfaces the
  // same signal as the `change` event carrying a fresh `IWorkbookData`. Skip
  // the echo from our own `persist()` setContent so a local edit doesn't wipe
  // the selection / re-read what we just wrote.
  useEffect(() => {
    const off = api.on('change', (snapshot) => {
      if (persistingRef.current) return;
      setCharts(readChartsFromSnapshot(snapshot));
      setSelectedId(null);
    });
    return off;
  }, [api]);

  const insert = useCallback(
    (chart: ChartModel) => {
      setCharts((prev) => {
        const next = [...prev, chart];
        notifyLocal(next);
        persist(next);
        return next;
      });
    },
    [notifyLocal, persist],
  );

  const remove = useCallback(
    (id: string) => {
      setCharts((prev) => {
        const next = prev.filter((c) => c.id !== id);
        notifyLocal(next);
        persist(next);
        return next;
      });
      setSelectedId((cur) => (cur === id ? null : cur));
    },
    [notifyLocal, persist],
  );

  const update = useCallback(
    (id: string, patch: Partial<ChartModel>) => {
      setCharts((prev) => {
        const next = prev.map((c) => (c.id === id ? { ...c, ...patch } : c));
        notifyLocal(next);
        persist(next);
        return next;
      });
    },
    [notifyLocal, persist],
  );

  const select = useCallback((id: string | null) => {
    setSelectedId(id);
  }, []);

  const __replaceAll = useCallback(
    (next: ChartModel[], opts?: { fromCollab?: boolean }) => {
      setCharts(next);
      // Drop the selection if its chart no longer exists in the new list
      // (a peer deleted it).
      setSelectedId((cur) => (cur && next.some((c) => c.id === cur) ? cur : null));
      if (!opts?.fromCollab) {
        notifyLocal(next);
        persist(next);
      }
    },
    [notifyLocal, persist],
  );

  const __subscribeLocal = useCallback((cb: (charts: ChartModel[]) => void) => {
    subsRef.current.add(cb);
    return () => {
      subsRef.current.delete(cb);
    };
  }, []);

  const value = useMemo<ChartsCtxValue>(
    () => ({
      api,
      charts,
      selectedId,
      insert,
      remove,
      update,
      select,
      __replaceAll,
      __subscribeLocal,
    }),
    [api, charts, selectedId, insert, remove, update, select, __replaceAll, __subscribeLocal],
  );

  return <ChartsContext.Provider value={value}>{children}</ChartsContext.Provider>;
}
