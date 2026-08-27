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
 * Pure scripting primitives behind CasualSheetsAPI's `executeCommands` /
 * `onMutation`. Kept free of any `@univerjs/*` *value* imports so it's unit
 * testable under the bare `node --import tsx` runner (importing Univer core
 * values fails to resolve there) — same split as
 * `read-only-predicate.ts` (pure) vs `read-only.ts` (wiring).
 */

/**
 * A single scriptable step — a Univer command/mutation id plus its params.
 * The unit a host records (via `onMutation`) and replays (via
 * `executeCommands`); the SDK generalization of the app's macro recorder.
 */
export interface CommandRecord {
  /** Univer command/mutation id, e.g. `sheet.mutation.set-range-values`. */
  id: string;
  /** The command's params object, passed straight back to `executeCommand`. */
  params?: object;
}

/**
 * Replay `steps` in order through `execute`. Best-effort: a step that throws or
 * resolves `false` (including a command-policy veto) is skipped. Resolves to the
 * number of steps the command bus accepted.
 */
export async function runSteps(
  execute: (id: string, params?: object) => Promise<unknown> | unknown,
  steps: CommandRecord[],
): Promise<number> {
  let applied = 0;
  for (const s of steps) {
    try {
      const result = await execute(s.id, s.params);
      if (result !== false) applied += 1;
    } catch {
      /* skip a step that no longer applies to the current state */
    }
  }
  return applied;
}

/** Minimal shape of the command service's collab mutation hook. */
export interface MutationEmitter {
  onMutationExecutedForCollab: (
    l: (info: CommandRecord, options?: MutationExecutionOptions) => void,
  ) => { dispose: () => void };
}

export interface MutationExecutionOptions {
  onlyLocal?: boolean;
  fromCollab?: boolean;
  fromChangeset?: boolean;
  syncOnly?: boolean;
}

type MutationObserverErrorHandler = (error: unknown) => void;

function notifyMutationObserver(
  handler: (record: CommandRecord) => void,
  record: CommandRecord,
  onError?: MutationObserverErrorHandler,
): void {
  try {
    handler(record);
  } catch (error) {
    // Mutation listeners run inside Univer's command stack after the state
    // change. A host exception must never turn a committed mutation into a
    // reported failure or strand that stack frame. Error reporting is itself
    // isolated so even a faulty reporter cannot escape the command bus.
    try {
      onError?.(error);
    } catch {
      /* observer failures are diagnostic only */
    }
  }
}

/**
 * Forward the collab mutation stream to `handler` as `{ id, params }` records.
 * Returns a disposer; safe to call with an absent service (no-op disposer).
 */
export function attachMutationObserver(
  cmdSvc: MutationEmitter | undefined,
  handler: (record: CommandRecord) => void,
  onError?: MutationObserverErrorHandler,
): () => void {
  const sub = cmdSvc?.onMutationExecutedForCollab((info) => {
    notifyMutationObserver(handler, { id: info.id, params: info.params }, onError);
  });
  return () => sub?.dispose();
}

/**
 * Observe only locally-authored, persistable mutations. Unlike the historical
 * `onMutation` audit stream, this excludes formula/cache-only writes plus
 * collaboration and changeset replay so a host can build a save queue without
 * echoing authoritative or derived state back to its server.
 */
export function attachLocalMutationObserver(
  cmdSvc: MutationEmitter | undefined,
  handler: (record: CommandRecord) => void,
  onError?: MutationObserverErrorHandler,
): () => void {
  const sub = cmdSvc?.onMutationExecutedForCollab((info, options) => {
    if (options?.onlyLocal || options?.fromCollab || options?.fromChangeset) return;
    notifyMutationObserver(handler, { id: info.id, params: info.params }, onError);
  });
  return () => sub?.dispose();
}
