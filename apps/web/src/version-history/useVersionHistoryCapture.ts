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

import { useEffect, useRef } from 'react';
import { ICommandService, type ICommandInfo, type IExecutionOptions } from '@univerjs/core';
import type { IWorkbookData } from '@univerjs/core';
import { useUniverAPI } from '../use-univer';
import { useWorkbook } from '../use-workbook';
import { useCollab } from '../collab/collab-context';
import { timeIt } from '../perf';
import { runWhenIdle, type IdleHandle } from '../idle';
import { createLiveVersionFeed, type LiveVersionFeed } from './live-feed';
import { setLiveFeed, writeVersion } from './store';

/**
 * Coarse-grained snapshot capture. Distinct from `useAutosave` which
 * runs a 5-second debounce into a single recovery slot. Version
 * history wants fewer, more deliberate captures:
 *
 *   - **Idle interval** ~10 min while the workbook is dirty since the
 *     last capture. Power users edit continuously; we don't want a
 *     snapshot per minute. Idle users get nothing — no point
 *     capturing identical state.
 *   - **Explicit save** via `saveNamedVersion(name)` exposed from
 *     this module (the File menu calls it). Manual snapshots are
 *     never pruned.
 *
 * Skipped inside co-edit rooms — the server keeps the authoritative
 * state and we'd just be duplicating local copies. Cross-peer version
 * history would need server-side support.
 *
 * The dirty flag is set by subscribing to `ICommandService.onMutationExecutedForCollab`
 * filtered to mutations that actually change document state (the same
 * filter `useAutosave` uses — selections and active-sheet swaps don't
 * count).
 */

const IDLE_INTERVAL_MS = 10 * 60 * 1000;

/**
 * Live feed singleton wired up at the same time the capture hook
 * mounts, so the panel always reads from the same feed instance the
 * store's notifier publishes to. Stored module-scope so other
 * consumers (e.g. the `saveNamedVersion` imperative API) can call
 * `feed.tick()` if they want optimistic UI without re-querying IDB.
 */
let liveFeed: LiveVersionFeed | null = null;
export function getLiveVersionFeed(): LiveVersionFeed {
  if (!liveFeed) {
    liveFeed = createLiveVersionFeed();
    setLiveFeed(liveFeed);
  }
  return liveFeed;
}

/** Capture a manual snapshot. Caller passes the name; this hook (or
 *  the file menu) handles the workbook readout + write. */
export async function saveNamedVersion(
  data: IWorkbookData,
  name: string,
  sourceFormat: string | null,
): Promise<number> {
  return writeVersion({
    kind: 'manual',
    name: name.trim() || 'Untitled version',
    savedAt: Date.now(),
    sourceFormat,
    data,
  });
}

export function useVersionHistoryCapture(): void {
  const api = useUniverAPI();
  const workbook = useWorkbook();
  const collab = useCollab();
  const dirtyRef = useRef(false);
  // Ensure the live feed exists before any subscribers mount.
  getLiveVersionFeed();

  useEffect(() => {
    if (!api) return;
    if (collab.roomId) return;

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const injector = (api as any)._injector as { get: (t: unknown) => unknown } | undefined;
    if (!injector) return;
    const cmdSvc = injector.get(ICommandService) as {
      onMutationExecutedForCollab: (
        l: (info: ICommandInfo, options?: IExecutionOptions) => void,
      ) => { dispose: () => void };
    };

    // Same "require user interaction first" guard the autosave hook
    // uses — Univer fires structural mutations during mount that
    // aren't real edits. Without this we'd capture a snapshot of the
    // empty default workbook on every fresh load.
    let userInteracted = false;
    const markInteracted = () => {
      userInteracted = true;
    };
    window.addEventListener('pointerdown', markInteracted, { capture: true });
    window.addEventListener('keydown', markInteracted, { capture: true });

    const sub = cmdSvc.onMutationExecutedForCollab((info, options) => {
      if (!userInteracted) return;
      if (options?.fromCollab) return;
      const id = info?.id ?? '';
      // Same skip list as the autosave hook — these mutations are
      // selection / view state, not document content.
      if (id.startsWith('sheet.mutation.set-selections')) return;
      if (id === 'sheet.mutation.set-worksheet-active-operation') return;
      dirtyRef.current = true;
    });

    let idle: IdleHandle | null = null;
    const capture = () => {
      if (!dirtyRef.current) return;
      const wb = api.getActiveWorkbook();
      if (!wb) return;
      // The snapshot is a full deep clone; defer it to an idle slot so it never
      // freezes the grid mid-edit (the interval fires regardless of activity).
      const data = timeIt('version-snapshot', () => wb.save() as unknown as IWorkbookData);
      void writeVersion({
        kind: 'auto',
        name: deriveAutoLabel(workbook.meta.name),
        savedAt: Date.now(),
        sourceFormat: workbook.meta.sourceFormat ?? null,
        data,
      })
        .then(() => {
          dirtyRef.current = false;
        })
        .catch((err) => console.warn('[version-history] auto-capture failed', err));
    };

    const tick = setInterval(() => {
      if (!dirtyRef.current) return;
      idle?.cancel();
      idle = runWhenIdle(capture);
    }, IDLE_INTERVAL_MS);

    return () => {
      sub.dispose();
      clearInterval(tick);
      idle?.cancel();
      window.removeEventListener('pointerdown', markInteracted, { capture: true });
      window.removeEventListener('keydown', markInteracted, { capture: true });
    };
    // workbook.meta is read fresh at each tick; not a dep to avoid
    // tearing down the interval on every name edit.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [api, collab.roomId]);
}

function deriveAutoLabel(workbookName: string): string {
  // Short label so the panel list reads cleanly. Time of day is
  // already rendered as a relative timestamp ("2 min ago"); the label
  // just answers "what was this snapshot of?".
  return `${workbookName || 'Workbook'} — auto`;
}
