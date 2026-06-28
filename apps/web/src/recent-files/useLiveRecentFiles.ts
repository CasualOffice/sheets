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
