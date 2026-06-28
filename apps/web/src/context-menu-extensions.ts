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

import type { Univer } from '@univerjs/core';
import { CommandType, ICommandService } from '@univerjs/core';
import {
  IMenuManagerService,
  MenuItemType,
  ContextMenuGroup,
  ContextMenuPosition,
} from '@univerjs/ui';

/**
 * Augments Univer's built-in cell context menu with operations we want there
 * but that the stock sheets-ui schema only puts on the ribbon — e.g. Merge —
 * plus thin proxy commands for our own dialogs (e.g. Paste Special).
 *
 * Univer doesn't expose its menu-item factories from `@univerjs/sheets-ui`
 * publicly, so we register proxy factories pointing at command ids. For
 * commands Univer already has (merge), we just register the menu entry.
 * For our dialog hooks (paste-special), we register a tiny custom command
 * here whose handler dispatches a DOM event the React shell listens for.
 */

const CASUAL_PASTE_SPECIAL_COMMAND_ID = 'casual.command.open-paste-special';
const CASUAL_FORMAT_CELLS_COMMAND_ID = 'casual.command.open-format-cells';

export function extendContextMenu(univer: Univer): void {
  const injector = univer.__getInjector();
  const menuMgr = injector.get(IMenuManagerService);
  const commandService = injector.get(ICommandService);

  // Register the open-paste-special command exactly once. The handler
  // dispatches a DOM event that the React MenuBar listens for and uses
  // to show <PasteSpecialDialog>. Going through a DOM event keeps the
  // Univer DI graph and React state cleanly separated — no shared store,
  // no facade-injection from React into Univer.
  if (!commandService.hasCommand(CASUAL_PASTE_SPECIAL_COMMAND_ID)) {
    commandService.registerCommand({
      id: CASUAL_PASTE_SPECIAL_COMMAND_ID,
      type: CommandType.OPERATION,
      handler: () => {
        document.dispatchEvent(new CustomEvent('casual-open-paste-special'));
        return true;
      },
    });
  }

  // Same DOM-event proxy pattern as Paste Special. MenuBar.tsx already
  // listens for `casual-open-format-cells` (registered alongside the
  // Ctrl+1 keyboard shortcut) and shows <FormatCellsDialog>. Excel's
  // right-click menu has "Format Cells…" as a top-level entry next to
  // Insert/Delete — we sit it in the FORMAT group so it groups with
  // the other formatting actions.
  if (!commandService.hasCommand(CASUAL_FORMAT_CELLS_COMMAND_ID)) {
    commandService.registerCommand({
      id: CASUAL_FORMAT_CELLS_COMMAND_ID,
      type: CommandType.OPERATION,
      handler: () => {
        document.dispatchEvent(new CustomEvent('casual-open-format-cells'));
        return true;
      },
    });
  }

  menuMgr.mergeMenu({
    [ContextMenuPosition.MAIN_AREA]: {
      [ContextMenuGroup.FORMAT]: {
        'sheet.command.add-worksheet-merge-all': {
          order: 100,
          menuItemFactory: () => ({
            id: 'sheet.command.add-worksheet-merge-all',
            type: MenuItemType.BUTTON,
            icon: 'MergeAllSingle',
            title: 'Merge cells',
          }),
        },
        'sheet.command.remove-worksheet-merge': {
          order: 101,
          menuItemFactory: () => ({
            id: 'sheet.command.remove-worksheet-merge',
            type: MenuItemType.BUTTON,
            icon: 'CancelMergeSingle',
            title: 'Unmerge',
          }),
        },
        [CASUAL_FORMAT_CELLS_COMMAND_ID]: {
          order: 110,
          menuItemFactory: () => ({
            id: CASUAL_FORMAT_CELLS_COMMAND_ID,
            type: MenuItemType.BUTTON,
            icon: 'BrushSingle',
            title: 'Format Cells…',
          }),
        },
      },
      // Add Paste Special to the QUICK group so it sits beside Cut /
      // Copy / Paste at the top of the menu — matches Excel's placement.
      [ContextMenuGroup.QUICK]: {
        [CASUAL_PASTE_SPECIAL_COMMAND_ID]: {
          order: 50,
          menuItemFactory: () => ({
            id: CASUAL_PASTE_SPECIAL_COMMAND_ID,
            type: MenuItemType.BUTTON,
            icon: 'PasteSpecial',
            title: 'Paste Special…',
          }),
        },
      },
    },
    // Mirror the cell-area Format Cells… entry into the row + column
    // header context menus. Excel right-clicks on a row or column
    // header also surface Format Cells…; without this the entry only
    // appears when right-clicking inside the grid.
    [ContextMenuPosition.ROW_HEADER]: {
      [ContextMenuGroup.FORMAT]: {
        [CASUAL_FORMAT_CELLS_COMMAND_ID]: {
          order: 110,
          menuItemFactory: () => ({
            id: CASUAL_FORMAT_CELLS_COMMAND_ID,
            type: MenuItemType.BUTTON,
            icon: 'BrushSingle',
            title: 'Format Cells…',
          }),
        },
      },
    },
    [ContextMenuPosition.COL_HEADER]: {
      [ContextMenuGroup.FORMAT]: {
        [CASUAL_FORMAT_CELLS_COMMAND_ID]: {
          order: 110,
          menuItemFactory: () => ({
            id: CASUAL_FORMAT_CELLS_COMMAND_ID,
            type: MenuItemType.BUTTON,
            icon: 'BrushSingle',
            title: 'Format Cells…',
          }),
        },
      },
    },
  });
}
