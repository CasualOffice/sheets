/**
 * ChromeTop — the chrome ABOVE the grid (menu bar + toolbar + formula bar).
 * Lives in the `@casualoffice/sheets/chrome` entry, lazy-imported by
 * `<CasualSheets>` only when `chrome !== 'none'` so bare-grid consumers (the
 * default; the apps/web host renders chrome="none" + its own shell) never bundle
 * the chrome JS.
 */
import { MenuBar } from './MenuBar';
import { Toolbar } from './Toolbar';
import { FormulaBar } from './FormulaBar';
import type { CasualSheetsAPI } from '../sheets/api';

export function ChromeTop({ api }: { api: CasualSheetsAPI | null }) {
  return (
    <>
      <MenuBar api={api} />
      <Toolbar api={api} />
      <FormulaBar api={api} />
    </>
  );
}
