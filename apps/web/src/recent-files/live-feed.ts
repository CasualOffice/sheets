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

/**
 * Tiny pub/sub that wakes recent-files list subscribers on every
 * write. Same shape as the version-history feed (see
 * apps/web/src/version-history/live-feed.ts) — single global
 * subscriber set, no granularity.
 */

export type LiveRecentFilesFeed = {
  tick: () => void;
  subscribe: (fn: () => void) => () => void;
};

export function createLiveRecentFilesFeed(): LiveRecentFilesFeed {
  const subs = new Set<() => void>();
  return {
    tick: () => {
      for (const fn of subs) {
        try {
          fn();
        } catch (err) {
          console.warn('[recent-files] subscriber threw', err);
        }
      }
    },
    subscribe: (fn) => {
      subs.add(fn);
      return () => {
        subs.delete(fn);
      };
    },
  };
}
