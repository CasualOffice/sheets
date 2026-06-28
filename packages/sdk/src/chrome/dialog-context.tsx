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
 * Dialog host + registry for `<CasualSheets chrome>`.
 *
 * The model: a chrome control (menu item, toolbar button) calls
 * `openDialog(kind, context)`. By default that opens the SDK's BUILT-IN dialog
 * for that kind (rendered by `<DialogHost>`). A host can OVERRIDE any kind — or
 * add novel kinds — via `extensions.dialogs`, and can additionally route
 * specific kinds to its legacy `onDialogRequest` callback (back-compat).
 *
 * Resolution order in `openDialog(kind)`:
 *   1. host registered `extensions.dialogs[kind]`        → render host dialog
 *   2. host opted `kind` to `onDialogRequest` (no built-in, or host wants it)
 *                                                        → emit onDialogRequest
 *   3. SDK has a built-in dialog for `kind`              → render built-in
 *   4. else, if `onDialogRequest` provided               → emit onDialogRequest
 *   5. else                                              → no-op
 *
 * `onDialogRequest` stays fully supported: if a host passes it and DOESN'T
 * register a built-in override, kinds the SDK has no built-in for fall through
 * to it (3→4), matching the previous behaviour.
 */

import { createContext, useCallback, useContext, useMemo, useState, type ReactNode } from 'react';
import type { CasualSheetsAPI } from '../sheets/api';
import type { ChromeExtensions, DialogComponentProps } from './extensions';
import { FormatCellsDialog } from './FormatCellsDialog';

/**
 * All dialog kinds the chrome knows about. Mirrors `MenuDialogKind` in MenuBar
 * (kept in sync) and is the key space for `extensions.dialogs`. Kinds with a
 * built-in renderer are listed in `BUILT_IN_DIALOGS`; the rest need a host
 * override or `onDialogRequest` to do anything.
 */
export type DialogKind =
  | 'format-cells'
  | 'find-replace'
  | 'insert-cells'
  | 'delete-cells'
  | 'paste-special'
  | 'insert-chart'
  | 'insert-pivot'
  | 'insert-sparkline'
  | 'insert-function'
  | 'name-manager'
  | 'goal-seek'
  | 'data-validation'
  | 'conditional-formatting'
  | 'custom-sort'
  | 'properties'
  | 'about'
  | 'keyboard-shortcuts';

/**
 * SDK built-in dialog components, keyed by kind. Anything NOT here has no
 * built-in UI and falls through to a host override / onDialogRequest.
 *
 * `find-replace` is intentionally NOT here — it's the self-managing
 * `<FindReplace>` panel (Ctrl+F/H), wired separately in ChromeTop and opened
 * imperatively; `DialogHost` only owns the modal-style built-ins.
 */
const BUILT_IN_DIALOGS: Partial<Record<DialogKind, React.ComponentType<DialogComponentProps>>> = {
  'format-cells': FormatCellsDialog,
};

/** Kinds the chrome can open without a host (built-in modal or self-managing). */
export function hasBuiltInDialog(kind: DialogKind): boolean {
  return kind in BUILT_IN_DIALOGS || kind === 'find-replace';
}

interface ActiveDialog {
  kind: DialogKind;
  context?: unknown;
}

export interface DialogController {
  /** Open a dialog (built-in by default; host override / onDialogRequest wins per rules). */
  openDialog: (kind: DialogKind, context?: unknown) => void;
  /** Close the currently-open built-in/override dialog. */
  closeDialog: () => void;
  /** True when the chrome can open `kind` itself OR a host registered/handles it. */
  canOpen: (kind: DialogKind) => boolean;
}

const DialogContext = createContext<DialogController | null>(null);

/** Access the dialog controller from any chrome component under the provider. */
export function useDialogs(): DialogController {
  const ctx = useContext(DialogContext);
  if (!ctx) {
    // Defensive — should never happen inside the chrome tree. A no-op keeps a
    // mis-mounted control from throwing.
    return {
      openDialog: () => {},
      closeDialog: () => {},
      canOpen: () => false,
    };
  }
  return ctx;
}

export interface DialogProviderProps {
  api: CasualSheetsAPI | null;
  extensions?: ChromeExtensions;
  /**
   * Legacy host hook. When set, kinds the host opts to handle (no built-in, or
   * explicitly listed in `hostOwnedDialogs`) emit through it instead of
   * rendering a built-in. Back-compat: still fires for non-built-in kinds.
   */
  onDialogRequest?: (kind: DialogKind, context?: unknown) => void;
  /**
   * Kinds the host wants to own via `onDialogRequest` even though the SDK has a
   * built-in. Lets a host keep the SDK chrome but its own Format Cells, say,
   * without registering a React component.
   */
  hostOwnedDialogs?: DialogKind[];
  /**
   * Bridge for the self-managing find-replace panel. `find-replace` isn't a
   * modal `<DialogHost>` renders — it's the `<FindReplace>` panel — so the
   * provider delegates opening it to the chrome (which owns the panel's open
   * signal) instead of setting `active`.
   */
  onOpenFindReplace?: (replaceMode: boolean) => void;
  children: ReactNode;
}

/**
 * Provides `openDialog`/`closeDialog` to the chrome tree and renders the active
 * built-in/override dialog via `<DialogHost>` (mounted internally).
 */
export function DialogProvider({
  api,
  extensions,
  onDialogRequest,
  hostOwnedDialogs,
  onOpenFindReplace,
  children,
}: DialogProviderProps) {
  const [active, setActive] = useState<ActiveDialog | null>(null);

  const overrides = extensions?.dialogs;
  const hostOwned = useMemo(() => new Set(hostOwnedDialogs ?? []), [hostOwnedDialogs]);

  const openDialog = useCallback(
    (kind: DialogKind, context?: unknown) => {
      // 1. Host-registered React override → render it.
      if (overrides?.[kind]) {
        setActive({ kind, context });
        return;
      }
      // 2. Host explicitly owns this kind via onDialogRequest.
      if (onDialogRequest && hostOwned.has(kind)) {
        onDialogRequest(kind, context);
        return;
      }
      // find-replace is the self-managing panel, not a DialogHost modal — bridge
      // to the chrome's open signal.
      if (kind === 'find-replace') {
        onOpenFindReplace?.(false);
        return;
      }
      // 3. SDK built-in.
      if (hasBuiltInDialog(kind)) {
        setActive({ kind, context });
        return;
      }
      // 4. Fall through to onDialogRequest for kinds with no built-in.
      if (onDialogRequest) {
        onDialogRequest(kind, context);
        return;
      }
      // 5. No way to fulfil — no-op (the control shouldn't have rendered).
    },
    [overrides, onDialogRequest, hostOwned, onOpenFindReplace],
  );

  const closeDialog = useCallback(() => setActive(null), []);

  const canOpen = useCallback(
    (kind: DialogKind) =>
      !!overrides?.[kind] || hasBuiltInDialog(kind) || !!onDialogRequest || hostOwned.has(kind),
    [overrides, onDialogRequest, hostOwned],
  );

  const controllerCanOpen = useCallback(
    (kind: DialogKind) =>
      kind === 'find-replace' ? !!onOpenFindReplace || canOpen(kind) : canOpen(kind),
    [canOpen, onOpenFindReplace],
  );

  const controller = useMemo<DialogController>(
    () => ({ openDialog, closeDialog, canOpen: controllerCanOpen }),
    [openDialog, closeDialog, controllerCanOpen],
  );

  return (
    <DialogContext.Provider value={controller}>
      {children}
      <DialogHost api={api} active={active} overrides={overrides} onClose={closeDialog} />
    </DialogContext.Provider>
  );
}

interface DialogHostProps {
  api: CasualSheetsAPI | null;
  active: ActiveDialog | null;
  overrides?: ChromeExtensions['dialogs'];
  onClose: () => void;
}

/**
 * Renders the currently-active modal dialog — a host override if registered for
 * the kind, otherwise the SDK built-in. `find-replace` is excluded (the
 * self-managing FindReplace panel owns it). Renders nothing when no dialog is
 * open or the API isn't ready.
 */
function DialogHost({ api, active, overrides, onClose }: DialogHostProps) {
  if (!active || !api) return null;
  if (active.kind === 'find-replace') return null;
  const Component = overrides?.[active.kind] ?? BUILT_IN_DIALOGS[active.kind];
  if (!Component) return null;
  return <Component api={api} onClose={onClose} context={active.context} />;
}
