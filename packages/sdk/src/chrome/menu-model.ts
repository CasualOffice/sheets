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
 * MenuBar data model + pure gating engine.
 *
 * Split out of `MenuBar.tsx` so the feature-gating contract is unit-testable
 * under `node --test`: `MenuBar.tsx` statically imports `@univerjs/core` values
 * (which the vendored typeless-package ESM can't expose as named exports to
 * node), so it can't be imported in a DOM-less test. This module has only
 * type-only imports (all erased at build), so `computeVisibleMenus` — the code
 * that decides which menus/items a host actually sees — can be exercised
 * directly.
 *
 * Feature gates: pass `features` to hide a control or whole menu group when its
 * feature is disabled. Defaults to all-enabled. A control whose feature is
 * `false` does not render. An entire top-level menu whose own `feature` is
 * `false`, or that ends up with no runnable items, is dropped.
 */

import type { DialogKind } from './dialog-context';
import type { CasualSheetsAPI } from '../sheets/api';
import type { MenuExtension } from './extensions';

export type MenuId = 'file' | 'edit' | 'view' | 'insert' | 'format' | 'data' | 'help';

/**
 * Dialog kinds the host can choose to render via `onDialogRequest`. These are
 * the actions the SDK chrome can't fulfil on its own (no built-in modal). The
 * string is passed straight to the host hook; the `context` (when present)
 * carries the pre-resolved A1 selection so the host doesn't have to re-read it.
 */
export type MenuDialogKind = DialogKind;

export type RunFn = (api: CasualSheetsAPI) => void;

export type MenuItemDef =
  | {
      kind: 'item';
      id: string;
      label: string;
      icon?: string;
      shortcut?: string;
      /** Dispatch a command / facade call directly. */
      run?: RunFn;
      /** Route through the host's `onDialogRequest`. Omitted if no host hook. */
      dialog?: MenuDialogKind;
      /** Feature gate — item hidden when `features[feature] === false`. */
      feature?: string;
    }
  | { kind: 'separator'; id: string; feature?: string }
  | {
      kind: 'submenu';
      id: string;
      label: string;
      icon?: string;
      items: MenuItemDef[];
      feature?: string;
    };

export interface MenuDef {
  id: MenuId;
  label: string;
  /** Feature gate for the whole menu. */
  feature?: string;
  items: MenuItemDef[];
}

/* ───────────────────────────── filtering ──────────────────────────────── */

/** True when the feature gate (if any) is enabled (default: enabled). */
export function featureOn(feature: string | undefined, features: Record<string, boolean>): boolean {
  if (!feature) return true;
  return features[feature] !== false;
}

/**
 * Keep an item if its feature is on AND — for a dialog item — the chrome can
 * open it (built-in dialog, host override, or `onDialogRequest`). Dialog items
 * with no way to open are dropped (the SDK never fakes a dialog). Submenus are
 * filtered recursively and dropped when empty.
 */
export function keepItem(
  item: MenuItemDef,
  features: Record<string, boolean>,
  canOpen: (kind: DialogKind) => boolean,
): MenuItemDef | null {
  if (!featureOn(item.feature, features)) return null;
  if (item.kind === 'separator') return item;
  if (item.kind === 'submenu') {
    const items = filterItems(item.items, features, canOpen);
    if (items.length === 0) return null;
    return { ...item, items };
  }
  if (item.dialog && !canOpen(item.dialog)) return null;
  return item;
}

/** Filter a list and collapse leading/trailing/double separators. */
export function filterItems(
  items: MenuItemDef[],
  features: Record<string, boolean>,
  canOpen: (kind: DialogKind) => boolean,
): MenuItemDef[] {
  const kept = items
    .map((i) => keepItem(i, features, canOpen))
    .filter((i): i is MenuItemDef => i !== null);
  // Collapse separators: drop leading, trailing, and runs.
  const out: MenuItemDef[] = [];
  for (const item of kept) {
    if (item.kind === 'separator') {
      if (out.length === 0) continue;
      if (out[out.length - 1].kind === 'separator') continue;
    }
    out.push(item);
  }
  while (out.length > 0 && out[out.length - 1].kind === 'separator') out.pop();
  return out;
}

/* ─────────────────────────── host extensions ──────────────────────────── */

/**
 * Append host menu extensions to their target top-level menu. Each extension
 * becomes a normal `item` (with a leading separator before the first host item
 * in that menu so it's visually grouped). Host items dispatch via `onClick` or
 * route a `dialog` kind through the dialog host, exactly like built-ins.
 */
export function withMenuExtensions(menus: MenuDef[], ext?: MenuExtension[]): MenuDef[] {
  if (!ext || ext.length === 0) return menus;
  const byMenu = new Map<MenuId, MenuExtension[]>();
  for (const e of ext) {
    const list = byMenu.get(e.menu) ?? [];
    list.push(e);
    byMenu.set(e.menu, list);
  }
  return menus.map((menu) => {
    const extras = byMenu.get(menu.id);
    if (!extras || extras.length === 0) return menu;
    const items: MenuItemDef[] = [...menu.items, { kind: 'separator', id: `ext-sep-${menu.id}` }];
    for (const e of extras) {
      items.push({
        kind: 'item',
        id: `ext-${e.id}`,
        label: e.label,
        icon: e.icon,
        shortcut: e.shortcut,
        dialog: e.dialog,
        run: e.onClick ? (api) => e.onClick?.(api) : undefined,
      });
    }
    return { ...menu, items };
  });
}

/**
 * Resolve the menus a host actually sees: append menu extensions, filter every
 * item by its `feature` gate + dialog-openability, then drop any top-level menu
 * whose own `feature` is off or that ends up empty.
 *
 * Exported (not inlined in the component) so the feature-gating contract — e.g.
 * `features={{ help: false }}` drops the Help menu and
 * `features={{ branding: false }}` drops the "View on GitHub" / "About" links —
 * is unit-testable without a DOM.
 */
export function computeVisibleMenus(
  menus: MenuDef[],
  features: Record<string, boolean>,
  canOpen: (kind: DialogKind) => boolean,
  ext?: MenuExtension[],
): MenuDef[] {
  return withMenuExtensions(menus, ext)
    .map((menu) => ({
      ...menu,
      items: filterItems(menu.items, features, canOpen),
    }))
    .filter((menu) => featureOn(menu.feature, features) && menu.items.length > 0);
}
