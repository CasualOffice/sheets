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

import { useEffect, useState } from 'react';
import { ICommandService, type ICommandInfo, type IExecutionOptions } from '@univerjs/core';
import type { FUniver } from '@univerjs/core/facade';
import { SetRangeValuesUndoMutationFactory } from '@univerjs/sheets';
import { REVERTABLE_MUTATIONS, SYNCED_MUTATIONS } from '@casualoffice/sheets/collab';

/**
 * Single-user version-history feed. The collab bridge already keeps a
 * shared op-log on the Yjs doc (see `bridge.ts`), but that source is
 * only present inside a room. For solo sessions we run a slimmed-down
 * version of the same capture path here:
 *
 *   - Subscribe to `ICommandService.onMutationExecutedForCollab`.
 *   - Filter to the same `SYNCED_MUTATIONS` set the bridge uses, so
 *     the panel surfaces match across modes.
 *   - For `REVERTABLE_MUTATIONS` (currently just set-range-values),
 *     capture undo params via Univer's existing factory in a
 *     `beforeCommandExecuted` pre-hook so the panel can revert.
 *   - Cap the in-memory ring at 500 entries — solo sessions don't have
 *     compaction; this is the upper bound a single user can scroll
 *     through usefully before opening Save / Undo instead.
 *
 * The shape is identical to the bridge's op-log entries so HistoryPanel
 * can render either source without branching on shape.
 */
export type LocalHistoryEntry = {
  c: string;
  t: number;
  id: string;
  p: unknown;
  u?: unknown;
};

const RING_CAP = 500;

export function useLocalHistory(api: FUniver | null): LocalHistoryEntry[] {
  const [entries, setEntries] = useState<LocalHistoryEntry[]>([]);

  useEffect(() => {
    if (!api) {
      setEntries([]);
      return;
    }
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const injector = (api as any)._injector as { get: (token: unknown) => unknown } | undefined;
    if (!injector) return;
    const cmdSvc = injector.get(ICommandService) as {
      onMutationExecutedForCollab: (
        l: (info: ICommandInfo, options?: IExecutionOptions) => void,
      ) => { dispose: () => void };
      beforeCommandExecuted: (l: (info: ICommandInfo, options?: IExecutionOptions) => void) => {
        dispose: () => void;
      };
    };

    // Same before-hook trick as the collab bridge: capture undo params
    // from pre-edit state, keyed by the params object so we can pair
    // them with the after-hook firing microseconds later.
    const pendingUndo = new Map<string, unknown>();
    const subBefore = cmdSvc.beforeCommandExecuted((info, options) => {
      if (options?.fromCollab) return;
      if (!REVERTABLE_MUTATIONS.has(info.id)) return;
      try {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const undo = SetRangeValuesUndoMutationFactory(injector as any, info.params as any);
        pendingUndo.set(JSON.stringify(info.params), undo);
      } catch {
        /* pre-edit state unreadable — entry just won't be revertable */
      }
    });

    const subAfter = cmdSvc.onMutationExecutedForCollab((info, options) => {
      if (options?.fromCollab) return;
      if (!SYNCED_MUTATIONS.has(info.id)) return;
      const key = JSON.stringify(info.params);
      const undoParams = pendingUndo.get(key);
      pendingUndo.delete(key);
      const entry: LocalHistoryEntry = {
        c: 'me',
        t: Date.now(),
        id: info.id,
        p: info.params,
        ...(undoParams !== undefined ? { u: undoParams } : {}),
      };
      setEntries((prev) => {
        const next = prev.length >= RING_CAP ? prev.slice(prev.length - RING_CAP + 1) : prev;
        return [...next, entry];
      });
    });

    return () => {
      subBefore.dispose();
      subAfter.dispose();
    };
  }, [api]);

  return entries;
}
