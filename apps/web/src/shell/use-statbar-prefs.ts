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

import { useSyncExternalStore } from 'react';

/**
 * Persistent preference for which selection stats appear in the
 * status bar. Mirrors Excel's right-click-on-status-bar checklist.
 * Stored via a small module-level store so two consumers (the stats
 * row + the customisation popover) always see the same flags.
 */

export type StatKey = 'avg' | 'count' | 'numCount' | 'min' | 'max' | 'sum';

export type StatPrefs = Record<StatKey, boolean>;

const STORAGE_KEY = 'casual:statbar-prefs';

const DEFAULTS: StatPrefs = {
  avg: true,
  count: true,
  numCount: true,
  min: true,
  max: true,
  sum: true,
};

function read(): StatPrefs {
  try {
    const raw = window.localStorage.getItem(STORAGE_KEY);
    if (!raw) return { ...DEFAULTS };
    const parsed = JSON.parse(raw) as Partial<StatPrefs>;
    return { ...DEFAULTS, ...parsed };
  } catch {
    return { ...DEFAULTS };
  }
}

let current: StatPrefs = read();
const subs = new Set<() => void>();

function subscribe(fn: () => void): () => void {
  subs.add(fn);
  return () => subs.delete(fn);
}
function snapshot(): StatPrefs {
  return current;
}

function setPrefs(next: StatPrefs): void {
  current = next;
  try {
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify(next));
  } catch {
    /* private mode — preference is in-memory only this session */
  }
  for (const fn of subs) {
    try {
      fn();
    } catch (err) {
      console.warn('[statbar-prefs] subscriber threw', err);
    }
  }
}

export function useStatPrefs(): {
  prefs: StatPrefs;
  toggle: (key: StatKey) => void;
} {
  const prefs = useSyncExternalStore(subscribe, snapshot, snapshot);
  const toggle = (key: StatKey) => setPrefs({ ...prefs, [key]: !prefs[key] });
  return { prefs, toggle };
}

export const STAT_LABELS: Record<StatKey, string> = {
  avg: 'Average',
  count: 'Count',
  numCount: 'Numerical Count',
  min: 'Min',
  max: 'Max',
  sum: 'Sum',
};
