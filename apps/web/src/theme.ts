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

import { useEffect, useSyncExternalStore } from 'react';
import { isDesktop } from './desk-bridge-bootstrap';

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

/**
 * Initial theme. Under the Casual Office desktop shell the launcher owns
 * the theme: the bridge bootstrap resolves `?theme=…` (incl. `system`)
 * and publishes the resolved value on `window.__deskApp__.theme` before
 * React mounts, so we seed from it and never read localStorage. The live
 * effect in `useTheme` keeps it in sync after mount. In a plain browser
 * `isDesktop()` is false and we keep the original localStorage behaviour.
 */
function readInitialTheme(): Theme {
  if (isDesktop()) {
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const t = (window as any).__deskApp__?.theme;
      if (t === 'dark' || t === 'light') return t;
    } catch {
      /* fall through to web behaviour */
    }
  }
  return readStoredTheme();
}

let currentTheme: Theme = readInitialTheme();
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

/**
 * Apply a theme pushed by the desktop launcher. Same store update +
 * `<html>` attribute + subscriber fan-out as `setTheme`, but it does NOT
 * persist to localStorage — the launcher is the source of truth in
 * desktop mode, so we must not overwrite the user's separate web choice.
 */
function applyExternalTheme(next: Theme): void {
  if (next === currentTheme) return;
  currentTheme = next;
  applyToHtml(next);
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

  // Desktop only: follow the launcher theme live. The bridge bootstrap
  // re-dispatches a `deskapp:theme` window CustomEvent (detail.resolved is
  // 'light'/'dark') on init, on Tauri `deskapp://theme` events, and on OS
  // scheme changes while tracking `system`. We mirror that into the store
  // without persisting, so `ThemeBridge` flips Univer in lockstep. Gated
  // behind `isDesktop()` so the web toggle/localStorage path is untouched.
  useEffect(() => {
    if (!isDesktop()) return;
    const onDeskTheme = (e: Event) => {
      const resolved = (e as CustomEvent<{ resolved?: Theme }>).detail?.resolved;
      if (resolved === 'dark' || resolved === 'light') applyExternalTheme(resolved);
    };
    window.addEventListener('deskapp:theme', onDeskTheme as EventListener);
    // Reconcile in case the bootstrap published before this effect ran.
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const t = (window as any).__deskApp__?.theme;
      if (t === 'dark' || t === 'light') applyExternalTheme(t);
    } catch {
      /* best-effort */
    }
    return () => window.removeEventListener('deskapp:theme', onDeskTheme as EventListener);
  }, []);

  return { theme, toggle };
}
