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
import { listVersions, type VersionSnapshot } from './store';
import { getLiveVersionFeed } from './useVersionHistoryCapture';

/**
 * Reactive snapshot list bound to the IDB version store. The store's
 * `notifyFeed` fires on every write / rename / delete; we re-query
 * IDB on each tick and update React state.
 *
 * Re-querying is cheap (the index is on `savedAt`, sorted in-memory
 * after `getAll`) and avoids tracking deltas in two places. If the
 * list ever grows into the thousands we'd switch to keying by id and
 * applying diff events.
 */
export function useLiveVersionList(): VersionSnapshot[] {
  const [list, setList] = useState<VersionSnapshot[]>([]);

  useEffect(() => {
    let cancelled = false;
    const refresh = () => {
      void listVersions().then((next) => {
        if (!cancelled) setList(next);
      });
    };
    refresh();
    const unsub = getLiveVersionFeed().subscribe(refresh);
    return () => {
      cancelled = true;
      unsub();
    };
  }, []);

  return list;
}
