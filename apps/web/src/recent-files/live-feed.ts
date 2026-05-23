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
