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
