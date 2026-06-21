/**
 * ChromeTop — the chrome ABOVE the grid (menu bar + toolbar + formula bar).
 * Lives in the `@casualoffice/sheets/chrome` entry, lazy-imported by
 * `<CasualSheets>` only when `chrome !== 'none'` so bare-grid consumers (the
 * default; the apps/web host renders chrome="none" + its own shell) never bundle
 * the chrome JS.
 *
 * Owns the dialog host: wraps MenuBar/Toolbar in `<DialogProvider>` so their
 * controls call `openDialog(kind)` and the SDK's built-in dialogs (Format Cells,
 * …) open by default — a host overrides them via `extensions.dialogs`. The
 * self-managing FindReplace panel is mounted here and bridged to
 * `openDialog('find-replace')` via an open signal.
 */
import { useCallback, useState } from 'react';
import { MenuBar } from './MenuBar';
import { Toolbar } from './Toolbar';
import { FormulaBar } from './FormulaBar';
import { FindReplace } from './FindReplace';
import { DialogProvider, type DialogKind } from './dialog-context';
import type { ChromeExtensions } from './extensions';
import type { CasualSheetsAPI } from '../sheets/api';

export interface ChromeTopProps {
  api: CasualSheetsAPI | null;
  /** Hide a control/group + block its command when its key is false. */
  features?: Record<string, boolean>;
  /**
   * Legacy host hook — route dialog-backed controls to the host instead of the
   * SDK built-in. Kept for back-compat; prefer `extensions.dialogs`.
   */
  onDialogRequest?: (kind: DialogKind, context?: unknown) => void;
  /** Kinds the host owns via `onDialogRequest` despite an SDK built-in. */
  hostOwnedDialogs?: DialogKind[];
  /** Host chrome extensions: custom toolbar/menu items, dialogs, panels. */
  extensions?: ChromeExtensions;
}

export function ChromeTop({
  api,
  features,
  onDialogRequest,
  hostOwnedDialogs,
  extensions,
}: ChromeTopProps) {
  // FindReplace is a self-managing panel; the dialog host opens it by bumping
  // this signal. (Ctrl/Cmd+F·H still open it directly.)
  const [findSignal, setFindSignal] = useState(0);
  const [findReplaceMode, setFindReplaceMode] = useState(false);
  const openFindReplace = useCallback((replaceMode: boolean) => {
    setFindReplaceMode(replaceMode);
    setFindSignal((n) => n + 1);
  }, []);

  return (
    <DialogProvider
      api={api}
      extensions={extensions}
      onDialogRequest={onDialogRequest}
      hostOwnedDialogs={hostOwnedDialogs}
      onOpenFindReplace={openFindReplace}
    >
      <MenuBar api={api} features={features} extensions={extensions} />
      <Toolbar api={api} features={features} extensions={extensions} />
      <FormulaBar api={api} />
      <FindReplace
        api={api}
        openSignal={findSignal || undefined}
        openInReplaceMode={findReplaceMode}
      />
    </DialogProvider>
  );
}
