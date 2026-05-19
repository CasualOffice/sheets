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
import { readPivotsFromSnapshot } from './resources';
import type { PivotModel } from './types';

/**
 * Pivot store mirrored into `IWorkbookData.resources['__casual_sheets_pivots__']`
 * at save time and re-hydrated when the active workbook changes. Same
 * shape as `ChartsContext`. The pivot's RENDERED cells live on the
 * sheet itself (a regular cell write); the model here is the definition
 * (source + target + field config) so a future refresh / change-source
 * action can re-run `applyPivot` without the user reconfiguring.
 */
type PivotsCtxValue = {
  pivots: PivotModel[];
  insert: (pivot: PivotModel) => void;
  remove: (id: string) => void;
  update: (id: string, patch: Partial<PivotModel>) => void;
};

export const PivotsContext = createContext<PivotsCtxValue>({
  pivots: [],
  insert: () => undefined,
  remove: () => undefined,
  update: () => undefined,
});

export function usePivots(): PivotsCtxValue {
  return useContext(PivotsContext);
}

export function PivotsProvider({ children }: { children: ReactNode }) {
  const api = useUniverAPI();
  const { meta, snapshotRef } = useWorkbook();
  const [pivots, setPivots] = useState<PivotModel[]>(() =>
    snapshotRef.current ? readPivotsFromSnapshot(snapshotRef.current) : [],
  );

  const lastRevisionRef = useRef(meta.revision);
  useEffect(() => {
    if (lastRevisionRef.current === meta.revision) return;
    lastRevisionRef.current = meta.revision;
    const snap = snapshotRef.current;
    if (!snap) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const wb = api?.getActiveWorkbook?.() as any;
      const fresh = wb?.save?.();
      setPivots(fresh ? readPivotsFromSnapshot(fresh) : []);
      return;
    }
    setPivots(readPivotsFromSnapshot(snap));
    // snapshotRef is a stable ref object — safe to exclude from deps.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [meta.revision, api]);

  const insert = useCallback((pivot: PivotModel) => {
    setPivots((prev) => [...prev, pivot]);
  }, []);

  const remove = useCallback((id: string) => {
    setPivots((prev) => prev.filter((p) => p.id !== id));
  }, []);

  const update = useCallback((id: string, patch: Partial<PivotModel>) => {
    setPivots((prev) => prev.map((p) => (p.id === id ? { ...p, ...patch } : p)));
  }, []);

  const value = useMemo<PivotsCtxValue>(
    () => ({ pivots, insert, remove, update }),
    [pivots, insert, remove, update],
  );

  return <PivotsContext.Provider value={value}>{children}</PivotsContext.Provider>;
}
