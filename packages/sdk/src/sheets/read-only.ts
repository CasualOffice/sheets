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

import { CustomCommandExecutionError, ICommandService, IPermissionService } from '@univerjs/core';
import type { FUniver } from '@univerjs/core/facade';
import { WorkbookEditablePermission } from '@univerjs/sheets';
import { isCommentOnlyBlocked, READONLY_BLOCK } from './read-only-predicate';

export { isCommentOnlyBlocked, isReadOnlyBlocked } from './read-only-predicate';

/**
 * Make a workbook genuinely READ-ONLY.
 *
 * Two layers, because they cover different host setups:
 *
 *  1. **Command veto** (`beforeCommandExecuted` → throw
 *     `CustomCommandExecutionError`) — cancels every mutating command. This is
 *     the load-bearing layer for the iframe embed, whose minimal plugin set does
 *     NOT enforce `WorkbookEditablePermission` (verified: the editor still
 *     accepts edits with the permission flipped off).
 *  2. **Permission flip** (`WorkbookEditablePermission` → false) — on the full
 *     `<CasualSheets>` host path this also greys out mutating menu items and
 *     stops the editor opening; harmless where unenforced.
 *
 * Returns a disposer that removes the veto and restores the prior editable
 * state (for callers that toggle a live unit between preview/editor).
 */
export function applyReadOnly(
  univerApi: FUniver,
  unitId: string,
  onBlock?: (commandId: string) => void,
): () => void {
  const injector = (univerApi as unknown as { _injector?: { get(t: unknown): unknown } })._injector;

  // Layer 1: veto mutating commands — the only layer the minimal embed enforces.
  const cmd = injector?.get(ICommandService) as
    | { beforeCommandExecuted(l: (info: { id: string }) => void): { dispose(): void } }
    | undefined;
  const vetoDisposable = cmd?.beforeCommandExecuted((info) => {
    if (READONLY_BLOCK.test(info.id)) {
      onBlock?.(info.id);
      throw new CustomCommandExecutionError(`read-only: blocked ${info.id}`);
    }
  });

  // Layer 2: permission flip (best-effort; load-bearing only on full hosts).
  const svc = injector?.get(IPermissionService) as
    | {
        addPermissionPoint(p: unknown): boolean;
        updatePermissionPoint(id: string, value: unknown): void;
        getPermissionPoint(id: string): { value: unknown } | undefined;
      }
    | undefined;

  // Constructing the point yields the deterministic id Univer derives for the
  // workbook-edit permission; we don't keep the instance otherwise.
  const id = new WorkbookEditablePermission(unitId).id;
  let prev: unknown;
  if (svc) {
    try {
      const existing = svc.getPermissionPoint(id);
      if (existing) {
        prev = existing.value;
        svc.updatePermissionPoint(id, false);
      } else {
        const point = new WorkbookEditablePermission(unitId);
        point.value = false;
        svc.addPermissionPoint(point);
      }
    } catch {
      /* best-effort — the veto above is the load-bearing layer */
    }
  }

  return () => {
    vetoDisposable?.dispose();
    try {
      svc?.updatePermissionPoint(id, prev === undefined ? true : prev);
    } catch {
      /* swallow */
    }
  };
}

/**
 * Make a workbook **comment-only** (share role `comment`): cells are read-only,
 * but threaded comments and their editor stay fully usable. This is the middle
 * tier between `view` (read-only, no comments) and `edit` (full), enforced at
 * the engine layer via the same command veto as {@link applyReadOnly} — minus
 * the comment commands.
 *
 * No `WorkbookEditablePermission` flip here: on a full host that point also
 * disables comment affordances, which would defeat the role. The command veto is
 * the load-bearing layer (and the only one the minimal embed enforces anyway).
 *
 * Returns a disposer that removes the veto.
 */
export function applyCommentOnly(
  univerApi: FUniver,
  onBlock?: (commandId: string) => void,
): () => void {
  const injector = (univerApi as unknown as { _injector?: { get(t: unknown): unknown } })._injector;
  const cmd = injector?.get(ICommandService) as
    | { beforeCommandExecuted(l: (info: { id: string }) => void): { dispose(): void } }
    | undefined;
  const vetoDisposable = cmd?.beforeCommandExecuted((info) => {
    if (isCommentOnlyBlocked(info.id)) {
      onBlock?.(info.id);
      throw new CustomCommandExecutionError(`comment-only: blocked ${info.id}`);
    }
  });
  return () => vetoDisposable?.dispose();
}

/**
 * Read the current `WorkbookEditablePermission` value for a unit — `true`
 * (editable), `false` (read-only), or `undefined` if the point isn't
 * registered yet. Lets hosts/tests confirm {@link applyReadOnly} took.
 */
export function getEditable(univerApi: FUniver, unitId: string): boolean | undefined {
  const injector = (univerApi as unknown as { _injector?: { get(t: unknown): unknown } })._injector;
  const svc = injector?.get(IPermissionService) as
    | { getPermissionPoint(id: string): { value: unknown } | undefined }
    | undefined;
  if (!svc) return undefined;
  const id = new WorkbookEditablePermission(unitId).id;
  const point = svc.getPermissionPoint(id);
  return point ? (point.value as boolean) : undefined;
}
