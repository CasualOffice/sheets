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
  type ReactNode,
} from 'react';
import { useUniverAPI } from '../use-univer';
import { useWorkbook } from '../use-workbook';
import { readSparklinesFromSnapshot } from './resources';
import type { SparklineModel } from './types';

/**
 * Store for the active workbook's sparklines. Mirrors `charts-context`
 * and `pivots-context` — hydrates from `IWorkbookData.resources` on
 * workbook swap, mirrors mutations back via the snapshot picked up by
 * autosave / xlsx export.
 */

type Ctx = {
  sparklines: SparklineModel[];
  add: (model: Omit<SparklineModel, 'id'>) => string;
  update: (id: string, patch: Partial<SparklineModel>) => void;
  remove: (id: string) => void;
  clear: () => void;
};

const SparklinesCtx = createContext<Ctx | null>(null);

export function SparklinesProvider({ children }: { children: ReactNode }) {
  const api = useUniverAPI();
  const { meta, snapshotRef } = useWorkbook();
  const [sparklines, setSparklines] = useState<SparklineModel[]>(() =>
    snapshotRef.current ? readSparklinesFromSnapshot(snapshotRef.current) : [],
  );

  // Re-hydrate on workbook swap. The snapshot ref is cleared shortly
  // after replaceWorkbook so we may need to fall back to wb.save().
  const lastRevisionRef = useRef(meta.revision);
  useEffect(() => {
    if (lastRevisionRef.current === meta.revision) return;
    lastRevisionRef.current = meta.revision;
    const snap = snapshotRef.current;
    if (!snap) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const wb = api?.getActiveWorkbook?.() as any;
      const fresh = wb?.save?.();
      setSparklines(fresh ? readSparklinesFromSnapshot(fresh) : []);
      return;
    }
    setSparklines(readSparklinesFromSnapshot(snap));
    // snapshotRef is a stable ref object — safe to exclude from deps.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [meta.revision, api]);

  const add = useCallback((model: Omit<SparklineModel, 'id'>): string => {
    const id = `spark-${Math.random().toString(36).slice(2, 10)}`;
    setSparklines((prev) => [...prev, { ...model, id }]);
    return id;
  }, []);

  const update = useCallback((id: string, patch: Partial<SparklineModel>) => {
    setSparklines((prev) => prev.map((s) => (s.id === id ? { ...s, ...patch } : s)));
  }, []);

  const remove = useCallback((id: string) => {
    setSparklines((prev) => prev.filter((s) => s.id !== id));
  }, []);

  const clear = useCallback(() => setSparklines([]), []);

  const value = useMemo<Ctx>(
    () => ({ sparklines, add, update, remove, clear }),
    [sparklines, add, update, remove, clear],
  );

  return <SparklinesCtx.Provider value={value}>{children}</SparklinesCtx.Provider>;
}

export function useSparklines(): Ctx {
  const ctx = useContext(SparklinesCtx);
  if (!ctx) throw new Error('useSparklines must be used inside <SparklinesProvider>');
  return ctx;
}
