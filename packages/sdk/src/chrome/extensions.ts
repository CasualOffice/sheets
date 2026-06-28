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
 * Chrome extension API — how a host builds its OWN chrome on top of the SDK's.
 *
 * `<CasualSheets chrome="full" extensions={...} />` lets a host APPEND custom
 * toolbar items / menu items / side panels and OVERRIDE (or add) dialogs by
 * kind. The SDK's built-ins are always the defaults; a host extension that
 * targets the same dialog `kind` wins. This is the "hosts must be able to build
 * their own on top of ours" contract — the SDK ships a working Office shell out
 * of the box, and a power host layers its product-specific bits without forking
 * the chrome.
 *
 * Every extension is handed the live `CasualSheetsAPI` (so it can read selection
 * / dispatch commands / reach `api.univer`), and dialog/panel components also
 * get an `onClose`.
 */

import type { ComponentType } from 'react';
import type { CasualSheetsAPI } from '../sheets/api';
import type { DialogKind } from './dialog-context';

/** Where on the menu strip a host menu item attaches. */
export type MenuTarget = 'file' | 'edit' | 'view' | 'insert' | 'format' | 'data' | 'help';

/** A host-supplied toolbar button, appended after the built-in toolbar groups. */
export interface ToolbarExtension {
  /** Stable id (used as React key + `data-testid` suffix `cs-ext-<id>`). */
  id: string;
  /** Accessible label (also the tooltip). */
  label: string;
  /** Material Symbols Outlined icon name (e.g. `'table_chart'`). */
  icon: string;
  /** Click handler. Receives the live API. Mutually exclusive with `command`. */
  onClick?: (api: CasualSheetsAPI) => void;
  /** Dispatch a Univer command id instead of an `onClick`. */
  command?: string;
  /** Optional command params (only with `command`). */
  commandParams?: object;
  /** Hide the item when this returns false (re-evaluated on command activity). */
  isVisible?: (api: CasualSheetsAPI) => boolean;
}

/** A host-supplied menu item, appended to the chosen top-level menu. */
export interface MenuExtension {
  /** Which top-level menu to append under. */
  menu: MenuTarget;
  /** Stable id (React key + `data-testid` `cs-menuitem-<id>`). */
  id: string;
  /** Item label. */
  label: string;
  /** Optional Material Symbols icon name. */
  icon?: string;
  /** Optional shortcut hint shown right-aligned (display only). */
  shortcut?: string;
  /** Click handler. Mutually exclusive with `dialog`. */
  onClick?: (api: CasualSheetsAPI) => void;
  /** Open a dialog by kind instead of an `onClick` (built-in or host-registered). */
  dialog?: DialogKind;
}

/** Props every dialog component (built-in or host override) receives. */
export interface DialogComponentProps {
  api: CasualSheetsAPI;
  /** Close the dialog. */
  onClose: () => void;
  /** Optional context passed at `openDialog(kind, context)` (e.g. A1 selection). */
  context?: unknown;
}

/** A host-supplied (or override) dialog, keyed by kind. */
export type DialogExtension = ComponentType<DialogComponentProps>;

/** Props every side-panel component receives. */
export interface PanelComponentProps {
  api: CasualSheetsAPI;
  /** Close/collapse the panel. */
  onClose: () => void;
}

/** A host-supplied side panel, surfaced on the panel rail. */
export interface PanelExtension {
  /** Stable id (React key + rail `data-testid` `cs-panel-<id>`). */
  id: string;
  /** Panel title shown in its header. */
  title: string;
  /** Material Symbols icon name for the rail button. */
  railIcon: string;
  /** The panel body. */
  component: ComponentType<PanelComponentProps>;
}

/**
 * The full chrome extension bundle a host passes to `<CasualSheets extensions>`.
 * Every field is optional — pass only what you extend.
 */
export interface ChromeExtensions {
  /** Custom toolbar buttons, appended after the built-in toolbar groups. */
  toolbar?: ToolbarExtension[];
  /** Custom menu items, appended to their chosen top-level menu. */
  menu?: MenuExtension[];
  /**
   * Dialogs by kind. A built-in kind here OVERRIDES the SDK's dialog; a novel
   * kind ADDS a host-only dialog (open it via a menu/toolbar item's `dialog`,
   * or imperatively through the API's `openDialog`).
   */
  dialogs?: Partial<Record<DialogKind, DialogExtension>>;
  /** Custom side panels, added to the panel rail. */
  panels?: PanelExtension[];
}

/** Empty default so consumers can destructure without null-checks. */
export const EMPTY_EXTENSIONS: ChromeExtensions = {};
