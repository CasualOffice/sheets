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

import { useEffect } from 'react';
import { ThemeService } from '@univerjs/core';
import { useUniverAPI } from '../use-univer';
import { useTheme } from '../theme';

/**
 * Mirror the React-side theme state into Univer's UI:
 *
 *   1. Toggle the `univer-dark` class on `<html>`. Univer's compiled
 *      CSS (`@univerjs/ui/lib/index.css`, `@univerjs/sheets-ui/lib/index.css`)
 *      ships dozens of `.univer-dark .univer-…` rules — they're the
 *      ONLY thing that flips column/row headers, gridlines, sidebar
 *      chrome, popovers, etc. to dark. Without this class on `<html>`
 *      Univer stays bright even if the React shell is dark.
 *
 *   2. Also flip `ThemeService.setDarkMode(true)`. Some Univer
 *      internals subscribe to `darkMode$` directly (notifications,
 *      message containers, mobile workbench) and skip the class-based
 *      path. We register both for belt-and-braces; the BehaviorSubject
 *      is idempotent on identical values.
 *
 * `Workbench.tsx` does this class toggle inside its own effect, but it
 * only runs when Univer renders its full UI. We configure
 * `UniverUIPlugin` with header/toolbar/footer disabled, so the root
 * Workbench DOES mount but a few of its layout effects can race our
 * theme change. Applying the class ourselves guarantees the dark CSS
 * wins regardless of which effect ran first.
 */
export function ThemeBridge() {
  const api = useUniverAPI();
  const { theme } = useTheme();

  // Class toggle — independent of api availability so the chrome is
  // dark from the first paint, before Univer mounts.
  useEffect(() => {
    const want = theme === 'dark';
    document.documentElement.classList.toggle('univer-dark', want);
  }, [theme]);

  // Univer-service flip — runs once api becomes available, and again
  // whenever the user toggles.
  useEffect(() => {
    if (!api) return;
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const injector = (api as any)._injector as
        | { get: (t: unknown) => unknown }
        | undefined;
      if (!injector) return;
      const themeService = injector.get(ThemeService) as
        | { setDarkMode: (b: boolean) => void; darkMode: boolean }
        | undefined;
      if (!themeService) return;
      const want = theme === 'dark';
      if (themeService.darkMode !== want) themeService.setDarkMode(want);
    } catch (err) {
      console.debug('[theme-bridge] could not reach Univer ThemeService', err);
    }
  }, [api, theme]);

  return null;
}
