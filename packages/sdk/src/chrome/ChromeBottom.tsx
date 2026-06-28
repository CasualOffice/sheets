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
 * ChromeBottom — the chrome BELOW the grid (sheet tabs + status bar).
 * Lazy-imported with ChromeTop (see ChromeTop).
 *
 * NOTE: Find & Replace is mounted in ChromeTop (inside the DialogProvider, so
 * the dialog host can open it via `openDialog('find-replace')`). It is NOT
 * mounted here — two `<FindReplace>` instances both bound Ctrl+F and rendered
 * two `cs-find-replace` dialogs (strict-mode violation in the chrome e2e).
 */
import { SheetTabs } from './SheetTabs';
import { StatusBar } from './StatusBar';
import type { CasualSheetsAPI } from '../sheets/api';

export function ChromeBottom({ api }: { api: CasualSheetsAPI | null }) {
  return (
    <>
      <SheetTabs api={api} />
      <StatusBar api={api} />
    </>
  );
}
