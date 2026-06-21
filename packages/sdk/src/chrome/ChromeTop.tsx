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

export interface ChromeTopProps {
  api: CasualSheetsAPI | null;
  /** Hide a control/group + block its command when its key is false. */
  features?: Record<string, boolean>;
  /** Route dialog-backed controls (Format Cells, Insert Chart, …) to the host. */
  onDialogRequest?: (kind: string, context?: unknown) => void;
}

export function ChromeTop({ api, features, onDialogRequest }: ChromeTopProps) {
  return (
    <>
      <MenuBar api={api} features={features} onDialogRequest={onDialogRequest} />
      <Toolbar api={api} features={features} onDialogRequest={onDialogRequest} />
      <FormulaBar api={api} />
    </>
  );
}
