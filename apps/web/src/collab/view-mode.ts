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
 * View-only enforcement for co-edit joiners with `role=view`.
 *
 * Previously the bridge filtered outbound mutations on the client
 * (`role === 'view' → return early in onMutationExecutedForCollab`),
 * but Univer's cell editor still opened on a double-click / F2 and
 * happily accepted keystrokes. The user would type, press Enter,
 * see the value locally, and nothing would sync — confusing and
 * impossible to debug.
 *
 * The right fix is to make Univer itself refuse the edit. Univer
 * 0.22 exposes a per-workbook `WorkbookEditablePermission` point;
 * when its `value` flips to `false`, the editor refuses to open and
 * the menu items that mutate cells go disabled. We reach into
 * IPermissionService via the FUniver injector and either add a fresh
 * point at value=false or flip an existing one.
 *
 * Returns a disposer that restores the previous value so role
 * changes during the session (rare; not exposed in the current UI)
 * don't strand the workbook in read-only forever.
 */
import type { FUniver } from '@univerjs/core/facade';
import { IPermissionService } from '@univerjs/core';
import { WorkbookEditablePermission } from '@univerjs/sheets';

export function applyViewOnlyMode(api: FUniver, unitId: string): () => void {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const injector = (api as any)._injector as
    | { get: (token: unknown) => unknown }
    | undefined;
  if (!injector) {
    console.warn('[collab][view-only] FUniver injector not accessible — read-only not enforced');
    return () => {};
  }
  let svc:
    | {
        addPermissionPoint: (p: unknown) => boolean;
        updatePermissionPoint: (id: string, value: unknown) => void;
        getPermissionPoint: (id: string) => { value: unknown } | undefined;
      }
    | undefined;
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    svc = injector.get(IPermissionService) as any;
  } catch (err) {
    console.warn('[collab][view-only] IPermissionService not available', err);
    return () => {};
  }
  if (!svc) return () => {};

  // Constructing the point just to read its id — Univer derives
  // `${UnitObject.Workbook}.${UnitAction.Edit}_${unitId}` deterministically.
  const point = new WorkbookEditablePermission(unitId);
  const id = point.id;
  let prev: unknown = undefined;
  try {
    const existing = svc.getPermissionPoint(id);
    if (existing) {
      prev = existing.value;
      svc.updatePermissionPoint(id, false);
    } else {
      point.value = false;
      svc.addPermissionPoint(point);
    }
  } catch (err) {
    console.warn('[collab][view-only] failed to apply read-only permission', err);
  }

  return () => {
    try {
      // Restore previous value if there was one (defaults true means the
      // workbook is editable again), otherwise leave the point alone —
      // the next workbook load will create fresh defaults.
      svc!.updatePermissionPoint(id, prev === undefined ? true : prev);
    } catch (err) {
      console.warn('[collab][view-only] failed to restore editable permission', err);
    }
  };
}
