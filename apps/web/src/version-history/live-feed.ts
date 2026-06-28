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
