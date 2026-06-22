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
