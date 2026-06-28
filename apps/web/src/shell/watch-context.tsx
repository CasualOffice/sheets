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

import { createContext, useCallback, useContext, useMemo, useState } from 'react';
import type { ReactNode } from 'react';
import { addWatches, removeWatch, type Watch } from './watch-model';

/**
 * Session-scoped store for the Watch Window. Watches live in memory for the
 * session (not persisted to the workbook yet — a follow-up); the panel reads
 * each watched cell's live value/formula from Univer.
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
  const [watches, setWatches] = useState<Watch[]>([]);
  const seqRef = useState(() => ({ n: 0 }))[0];

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
