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

import { useUniverAPI } from '../use-univer';
import { useOutline } from './outline-context';
import type { OutlineGroup } from './types';

type Axis = 'rows' | 'cols';

/**
 * Hook over OutlineContext + Univer Facade that exposes the high-level
 * "Group / Ungroup the current selection" operations the menu binds to.
 * Lives in `outline/` because it owns the cross-cutting workflow; the
 * context itself stays focused on state.
 */
export function useOutlineActions() {
  const api = useUniverAPI();
  const outline = useOutline();

  const readActive = () => {
    if (!api) return null;
    const wb = api.getActiveWorkbook();
    if (!wb) return null;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = wb.getActiveSheet();
    if (!ws) return null;
    const sheetId = typeof ws.getSheetId === 'function' ? ws.getSheetId() : null;
    const range = typeof ws.getActiveRange === 'function' ? ws.getActiveRange() : null;
    if (!sheetId || !range) return null;
    return {
      sheetId: sheetId as string,
      startRow: range.getRow() as number,
      endRow: (range.getRow() as number) + (range.getHeight() as number) - 1,
      startCol: range.getColumn() as number,
      endCol: (range.getColumn() as number) + (range.getWidth() as number) - 1,
    };
  };

  const groupRows = () => {
    const a = readActive();
    if (!a) return null;
    return outline.addGroup(a.sheetId, 'rows', a.startRow, a.endRow);
  };

  const groupCols = () => {
    const a = readActive();
    if (!a) return null;
    return outline.addGroup(a.sheetId, 'cols', a.startCol, a.endCol);
  };

  /**
   * Remove the group(s) that contain the active cell — covers both row and
   * column groups in one click, which is what users expect from "Ungroup".
   */
  const ungroupSelection = (): { removed: number } => {
    const a = readActive();
    if (!a) return { removed: 0 };
    let removed = 0;
    const sheet = outline.getSheet(a.sheetId);
    const visit = (axis: Axis, anchor: number) => {
      const hit = sheet[axis].find((g) => anchor >= g.start && anchor <= g.end);
      if (hit) {
        outline.removeGroup(a.sheetId, axis, hit.id);
        removed++;
      }
    };
    visit('rows', a.startRow);
    visit('cols', a.startCol);
    return { removed };
  };

  const toggleCollapsed = (axis: Axis, group: OutlineGroup) => {
    const a = readActive();
    if (!a) return;
    outline.setCollapsed(a.sheetId, axis, group.id, !group.collapsed);
  };

  return { groupRows, groupCols, ungroupSelection, toggleCollapsed };
}
