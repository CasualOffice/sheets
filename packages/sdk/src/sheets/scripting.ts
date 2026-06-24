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
 * Replay `steps` in order through `execute`. Best-effort: a step that throws is
 * skipped (the underlying state may have moved on). Resolves to the number of
 * steps that ran without throwing.
 */
export async function runSteps(
  execute: (id: string, params?: object) => Promise<unknown> | unknown,
  steps: CommandRecord[],
): Promise<number> {
  let applied = 0;
  for (const s of steps) {
    try {
      await execute(s.id, s.params);
      applied += 1;
    } catch {
      /* skip a step that no longer applies to the current state */
    }
  }
  return applied;
}

/** Minimal shape of the command service's collab mutation hook. */
export interface MutationEmitter {
  onMutationExecutedForCollab: (l: (info: CommandRecord) => void) => { dispose: () => void };
}

/**
 * Forward the collab mutation stream to `handler` as `{ id, params }` records.
 * Returns a disposer; safe to call with an absent service (no-op disposer).
 */
export function attachMutationObserver(
  cmdSvc: MutationEmitter | undefined,
  handler: (record: CommandRecord) => void,
): () => void {
  const sub = cmdSvc?.onMutationExecutedForCollab((info) =>
    handler({ id: info.id, params: info.params }),
  );
  return () => sub?.dispose();
}
