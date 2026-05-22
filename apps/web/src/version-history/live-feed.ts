/**
 * Tiny pub/sub the version-history store uses to wake subscribers
 * whenever a snapshot is written / renamed / deleted. The panel
 * subscribes via `useLiveVersionList` and rebuilds its list on tick.
 *
 * Kept dead simple — one global subscription set, no granularity, no
 * unsubscribe leaks (the panel pairs subscribe/unsubscribe in the
 * effect cleanup). For a single-user app this is enough; if we ever
 * had thousands of versions we'd switch to a diff-based event.
 */
export type LiveVersionFeed = {
  tick: () => void;
  subscribe: (fn: () => void) => () => void;
};

export function createLiveVersionFeed(): LiveVersionFeed {
  const subs = new Set<() => void>();
  return {
    tick: () => {
      for (const fn of subs) {
        try {
          fn();
        } catch (err) {
          console.warn('[version-history] subscriber threw', err);
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
