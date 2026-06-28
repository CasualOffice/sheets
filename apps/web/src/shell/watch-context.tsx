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
import { useWorkbook } from '../use-workbook';
import { addWatches, removeWatch, type Watch } from './watch-model';
import { readWatchesFromSnapshot } from './watch-resources';

/**
 * Store for the Watch Window. Mirrored into
 * `IWorkbookData.resources['__casual_sheets_watches__']` at xlsx export and
 * re-hydrated when the active workbook changes (same pattern as pivots/charts;
 * autosave's `wb.save()` snapshot doesn't carry it). The panel reads each
 * watched cell's live value/formula from Univer.
 */
type WatchCtxValue = {
  watches: Watch[];
  add: (cells: Array<Omit<Watch, 'id'>>) => void;
  remove: (id: string) => void;
  clear: () => void;
};

const WatchContext = createContext<WatchCtxValue | null>(null);

export function useWatches(): WatchCtxValue {
  const ctx = useContext(WatchContext);
  if (!ctx) throw new Error('useWatches must be used inside <WatchProvider>');
  return ctx;
}

export function WatchProvider({ children }: { children: ReactNode }) {
  const { meta, snapshotRef } = useWorkbook();
  const [watches, setWatches] = useState<Watch[]>(() =>
    snapshotRef.current ? readWatchesFromSnapshot(snapshotRef.current) : [],
  );
  const seqRef = useState(() => ({ n: 0 }))[0];

  // Re-hydrate from the snapshot when the active workbook changes (open / swap).
  const lastRevisionRef = useRef(meta.revision);
  useEffect(() => {
    if (lastRevisionRef.current === meta.revision) return;
    lastRevisionRef.current = meta.revision;
    setWatches(snapshotRef.current ? readWatchesFromSnapshot(snapshotRef.current) : []);
    // snapshotRef is a stable ref object — safe to exclude.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [meta.revision]);

  const add = useCallback(
    (cells: Array<Omit<Watch, 'id'>>) => {
      setWatches((cur) => addWatches(cur, cells, () => `watch-${(seqRef.n += 1)}`));
    },
    [seqRef],
  );
  const remove = useCallback((id: string) => setWatches((cur) => removeWatch(cur, id)), []);
  const clear = useCallback(() => setWatches([]), []);

  const value = useMemo<WatchCtxValue>(
    () => ({ watches, add, remove, clear }),
    [watches, add, remove, clear],
  );
  return <WatchContext.Provider value={value}>{children}</WatchContext.Provider>;
}
