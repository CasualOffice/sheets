import { useSyncExternalStore } from 'react';

/**
 * Light/dark theme — single source of truth across every `useTheme()`
 * caller. The earlier draft used `useState` inside the hook, which
 * gave each consumer (TitleBar's toggle, ThemeBridge's reader) its
 * OWN independent copy of the flag. Toggling in one component never
 * propagated; the canvas stayed bright while the title bar flipped to
 * dark and back.
 *
 * `useSyncExternalStore` reads from this module-level state and
 * subscribes to a tiny pub/sub. All consumers see the same value,
 * re-render together when it changes, and the persisted choice is
 * loaded once at module init.
 *
 * Manual-only by design (no `prefers-color-scheme` subscription). The
 * `data-theme` attribute on `<html>` drives our chrome CSS; the
 * `univer-dark` class (applied by `ThemeBridge`) drives Univer's
 * canvas CSS.
 */

const STORAGE_KEY = 'casual:theme';

export type Theme = 'light' | 'dark';

function readStoredTheme(): Theme {
  try {
    const v = window.localStorage.getItem(STORAGE_KEY);
    if (v === 'dark' || v === 'light') return v;
  } catch {
    /* private mode etc. — fall through */
  }
  return 'light';
}

let currentTheme: Theme = readStoredTheme();
const subscribers = new Set<() => void>();

function applyToHtml(theme: Theme): void {
  document.documentElement.setAttribute('data-theme', theme);
}

// Apply once at module load so the very first paint is in the right
// scheme. Without this, dark users would see a flash of light chrome
// during React mount.
applyToHtml(currentTheme);

function setTheme(next: Theme): void {
  if (next === currentTheme) return;
  currentTheme = next;
  applyToHtml(next);
  try {
    window.localStorage.setItem(STORAGE_KEY, next);
  } catch {
    /* persistence is best-effort */
  }
  for (const fn of subscribers) {
    try {
      fn();
    } catch (err) {
      console.warn('[theme] subscriber threw', err);
    }
  }
}

function subscribe(fn: () => void): () => void {
  subscribers.add(fn);
  return () => {
    subscribers.delete(fn);
  };
}

function getSnapshot(): Theme {
  return currentTheme;
}

// `useSyncExternalStore` is the React-blessed way to wire a
// component to module-level state. SSR is irrelevant for us, but the
// `getServerSnapshot` arg is required — return the same value.
export function useTheme(): { theme: Theme; toggle: () => void } {
  const theme = useSyncExternalStore(subscribe, getSnapshot, getSnapshot);
  const toggle = () => setTheme(theme === 'dark' ? 'light' : 'dark');
  return { theme, toggle };
}
