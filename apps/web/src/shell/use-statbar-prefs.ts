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
