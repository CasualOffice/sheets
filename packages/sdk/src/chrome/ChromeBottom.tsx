/**
 * ChromeBottom — the chrome BELOW the grid (sheet tabs + status bar) plus the
 * Find & Replace overlay. Lazy-imported with ChromeTop (see ChromeTop).
 */
import { SheetTabs } from './SheetTabs';
import { StatusBar } from './StatusBar';
import { FindReplace } from './FindReplace';
import type { CasualSheetsAPI } from '../sheets/api';

export function ChromeBottom({ api }: { api: CasualSheetsAPI | null }) {
  return (
    <>
      <SheetTabs api={api} />
      <StatusBar api={api} />
      <FindReplace api={api} />
    </>
  );
}
