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

import { useEffect, useState } from 'react';
import { useFileSource } from './context';
import type { RecentEntry } from './types';

/**
 * Reactive recent-files list scoped to the active `FileSource`.
 *
 * Mirrors the older `useLiveRecentFiles` hook but routes through the
 * source so Mode 3 and Mode 2 can swap in their own listing (server
 * fetch + push notification) without touching this hook or the
 * components that consume it.
 *
 * Resubscribes when the source identity changes — Phase C will swap
 * the source on auth changes, and consumers should see the new list
 * without a remount.
 */

export function useRecentFiles(): RecentEntry[] {
  const source = useFileSource();
  const [list, setList] = useState<RecentEntry[]>([]);

  useEffect(() => {
    let cancelled = false;
    const refresh = () => {
      void source.listRecent().then((next) => {
        if (!cancelled) setList(next);
      });
    };
    refresh();
    const unsub = source.subscribeRecent(refresh);
    return () => {
      cancelled = true;
      unsub();
    };
  }, [source]);

  return list;
}
