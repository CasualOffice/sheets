import { createContext, useCallback, useContext, useMemo, useState } from 'react';
import type { ReactNode } from 'react';
import type { ChartModel } from './types';

/**
 * In-memory chart store. P0 — non-persistent; P1 mirrors into
 * `IWorkbookData.resources['CASUAL_CHARTS']` so the model round-
 * trips through xlsx + collab via the existing hidden-sheet
 * resource mechanism. Splitting persistence out of P0 lets us
 * validate the rendering + insertion flow without committing to
 * the resource schema yet.
 *
 * The store keeps charts keyed by id (`ChartModel.id`). Removal
 * by id; updates by replace. No mutation in place — every change
 * produces a new model object so React's effect deps work.
 */
type ChartsCtxValue = {
  charts: ChartModel[];
  insert: (chart: ChartModel) => void;
  remove: (id: string) => void;
  update: (id: string, patch: Partial<ChartModel>) => void;
};

export const ChartsContext = createContext<ChartsCtxValue>({
  charts: [],
  insert: () => undefined,
  remove: () => undefined,
  update: () => undefined,
});

export function useCharts(): ChartsCtxValue {
  return useContext(ChartsContext);
}

export function ChartsProvider({ children }: { children: ReactNode }) {
  const [charts, setCharts] = useState<ChartModel[]>([]);

  const insert = useCallback((chart: ChartModel) => {
    setCharts((prev) => [...prev, chart]);
  }, []);

  const remove = useCallback((id: string) => {
    setCharts((prev) => prev.filter((c) => c.id !== id));
  }, []);

  const update = useCallback((id: string, patch: Partial<ChartModel>) => {
    setCharts((prev) => prev.map((c) => (c.id === id ? { ...c, ...patch } : c)));
  }, []);

  const value = useMemo<ChartsCtxValue>(
    () => ({ charts, insert, remove, update }),
    [charts, insert, remove, update],
  );

  return <ChartsContext.Provider value={value}>{children}</ChartsContext.Provider>;
}
