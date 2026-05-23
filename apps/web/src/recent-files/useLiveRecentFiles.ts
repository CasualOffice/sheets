import { useEffect, useState } from 'react';
import { listRecentFiles, setLiveFeed, type RecentFile } from './store';
import { createLiveRecentFilesFeed, type LiveRecentFilesFeed } from './live-feed';

/**
 * Reactive list of recent files. Each write to the IDB store notifies
 * via the shared `LiveRecentFilesFeed`; we re-query on every tick.
 *
 * Initialised lazily so the feed and IDB don't open until something
 * actually needs the list.
 */

let feed: LiveRecentFilesFeed | null = null;
function getFeed(): LiveRecentFilesFeed {
  if (!feed) {
    feed = createLiveRecentFilesFeed();
    setLiveFeed(feed);
  }
  return feed;
}

export function useLiveRecentFiles(): RecentFile[] {
  const [list, setList] = useState<RecentFile[]>([]);

  useEffect(() => {
    let cancelled = false;
    const refresh = () => {
      void listRecentFiles().then((next) => {
        if (!cancelled) setList(next);
      });
    };
    refresh();
    const unsub = getFeed().subscribe(refresh);
    return () => {
      cancelled = true;
      unsub();
    };
  }, []);

  return list;
}
