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
import { useUniverAPI } from '../use-univer';
import { useWorkbook } from '../use-workbook';
import { readChartsFromSnapshot } from './resources';
import type { ChartModel } from './types';

/**
 * Chart store mirrored into `IWorkbookData.resources['__casual_sheets_charts__']`
 * at save time and re-hydrated when the active workbook changes. The store
 * keeps charts keyed by id (`ChartModel.id`); removal by id, updates by
 * replace — every change produces a new model object so React effect deps
 * work for downstream consumers (ChartLayer, ChartOverlay).
 */
type ChartsCtxValue = {
  charts: ChartModel[];
  /** Id of the chart with the selection frame + handles drawn. At most
   *  one chart is selected at a time (Excel single-select; multi-select
   *  via Ctrl+click is not implemented yet). */
  selectedId: string | null;
  insert: (chart: ChartModel) => void;
  remove: (id: string) => void;
  update: (id: string, patch: Partial<ChartModel>) => void;
  select: (id: string | null) => void;
};

export const ChartsContext = createContext<ChartsCtxValue>({
  charts: [],
  selectedId: null,
  insert: () => undefined,
  remove: () => undefined,
  update: () => undefined,
  select: () => undefined,
});

export function useCharts(): ChartsCtxValue {
  return useContext(ChartsContext);
}

export function ChartsProvider({ children }: { children: ReactNode }) {
  const api = useUniverAPI();
  const { meta, snapshotRef } = useWorkbook();
  const [charts, setCharts] = useState<ChartModel[]>(() =>
    snapshotRef.current ? readChartsFromSnapshot(snapshotRef.current) : [],
  );
  const [selectedId, setSelectedId] = useState<string | null>(null);

  // Re-hydrate every time the workbook is replaced (Open / New / collab
  // remote-snapshot). Mirrors the outline-context rehydrate path — the
  // snapshot lives on `snapshotRef.current` only for the brief window
  // between replaceWorkbook and the post-render flush, which is exactly
  // the window we need to read it from here. Fall back to `wb.save()`
  // (deep clone) on rare misses where the ref was already cleared.
  const lastRevisionRef = useRef(meta.revision);
  useEffect(() => {
    if (lastRevisionRef.current === meta.revision) return;
    lastRevisionRef.current = meta.revision;
    const snap = snapshotRef.current;
    if (!snap) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const wb = api?.getActiveWorkbook?.() as any;
      const fresh = wb?.save?.();
      setCharts(fresh ? readChartsFromSnapshot(fresh) : []);
      return;
    }
    setCharts(readChartsFromSnapshot(snap));
    setSelectedId(null);
    // snapshotRef is a stable ref object — safe to exclude from deps.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [meta.revision, api]);

  const insert = useCallback((chart: ChartModel) => {
    setCharts((prev) => [...prev, chart]);
  }, []);

  const remove = useCallback((id: string) => {
    setCharts((prev) => prev.filter((c) => c.id !== id));
    setSelectedId((cur) => (cur === id ? null : cur));
  }, []);

  const update = useCallback((id: string, patch: Partial<ChartModel>) => {
    setCharts((prev) => prev.map((c) => (c.id === id ? { ...c, ...patch } : c)));
  }, []);

  const select = useCallback((id: string | null) => {
    setSelectedId(id);
  }, []);

  const value = useMemo<ChartsCtxValue>(
    () => ({ charts, selectedId, insert, remove, update, select }),
    [charts, selectedId, insert, remove, update, select],
  );

  return <ChartsContext.Provider value={value}>{children}</ChartsContext.Provider>;
}
