import * as Y from 'yjs';
import type { FUniver } from '@univerjs/core/facade';
import {
  ICommandService,
  type ICommandInfo,
  type IExecutionOptions,
} from '@univerjs/core';

/**
 * Yjs ↔ Univer mutation bridge. See docs/CO-EDITING.md for the design.
 *
 * Strategy: every non-collab mutation gets serialized into a Y.Array log;
 * peers replay each entry with `fromCollab: true`. The log is the source
 * of truth — late joiners read the whole array on connect via Yjs sync
 * and replay it once, ending up at the same state.
 *
 * Why an op log (not a state mirror): writing a per-mutation state mirror
 * for every mutation Univer emits (set-range-values, set-style, insert-row,
 * merge, hide-col, freeze, …) is dozens of handlers. The log generalizes
 * — any deterministic mutation just round-trips its params. Trade-off: no
 * per-cell CRDT merging on concurrent writes (Yjs orders inserts, then
 * Univer re-executes them; last writer wins at the mutation level). For
 * v1 that matches expectations.
 *
 * Echo-loop guard (per CLAUDE.md):
 *   - Records carry the emitter's `clientId`. The observer skips records
 *     it emitted itself (otherwise we'd double-apply our own writes).
 *   - Remote applies pass `fromCollab: true` so Univer's
 *     `onMutationExecutedForCollab` listener filters them back out via the
 *     options check below.
 */

const LOG_KEY = 'ops';

/**
 * Allowlist of mutation ids we sync. Listed explicitly to keep
 * undocumented / version-volatile mutations out of the log — anything
 * not here just stays local. Easier to add new ids than to debug a
 * silent corruption from a mutation that secretly references local
 * state (selections, render skeletons, etc.).
 */
const SYNCED_MUTATIONS: ReadonlySet<string> = new Set([
  // Cell-level — values, formulas, styles, rich text.
  'sheet.mutation.set-range-values',
  'sheet.mutation.set-style',
  // Row / column structural.
  'sheet.mutation.insert-row',
  'sheet.mutation.insert-col',
  'sheet.mutation.remove-row',
  'sheet.mutation.remove-col',
  'sheet.mutation.move-rows',
  'sheet.mutation.move-cols',
  'sheet.mutation.set-row-hidden',
  'sheet.mutation.set-row-visible',
  'sheet.mutation.set-col-hidden',
  'sheet.mutation.set-col-visible',
  'sheet.mutation.set-worksheet-row-height',
  'sheet.mutation.set-worksheet-row-is-auto-height',
  'sheet.mutation.set-worksheet-col-width',
  // Merges.
  'sheet.mutation.add-worksheet-merge',
  'sheet.mutation.remove-worksheet-merge',
  // Sheet lifecycle.
  'sheet.mutation.insert-sheet',
  'sheet.mutation.remove-sheet',
  'sheet.mutation.set-worksheet-name',
  'sheet.mutation.set-worksheet-order',
  // Freeze.
  'sheet.mutation.set-frozen',
  // Hyperlinks (sheets-hyper-link).
  'sheet.mutation.add-hyper-link',
  'sheet.mutation.remove-hyper-link',
  'sheet.mutation.update-hyper-link',
]);

type OpRecord = {
  /** Yjs client id of the emitter (string for portability via JSON). */
  c: string;
  /** Wall-clock at emit; diagnostic only. */
  t: number;
  /** Mutation id (e.g. `sheet.mutation.set-range-values`). */
  id: string;
  /** Mutation params, JSON-serializable. */
  p: unknown;
};

export type BridgeHandle = {
  /** Underlying Yjs document — exposed so tests / devtools can introspect. */
  doc: Y.Doc;
  /** Stop listening and detach. */
  dispose: () => void;
};

export type BridgeOptions = {
  /**
   * `view` clients only RECEIVE remote updates — local mutations don't
   * append to the log, so peers never see them. Client-side gate only;
   * a determined user can run the bridge in write mode by editing the
   * URL. Server-side enforcement is a follow-up hardening pass.
   */
  role?: 'view' | 'write';
};

export function startBridge(api: FUniver, doc: Y.Doc, opts: BridgeOptions = {}): BridgeHandle {
  const role = opts.role ?? 'write';
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const injector = (api as any)._injector as
    | { get: (token: unknown) => unknown }
    | undefined;
  if (!injector) {
    throw new Error('[collab] FUniver injector not accessible — Univer too old?');
  }
  const cmdSvc = injector.get(ICommandService) as {
    onMutationExecutedForCollab: (
      l: (info: ICommandInfo, options?: IExecutionOptions) => void,
    ) => { dispose: () => void };
    executeCommand: (id: string, params: unknown, options?: IExecutionOptions) => Promise<unknown>;
  };

  const log = doc.getArray<OpRecord>(LOG_KEY);
  const myClientId = String(doc.clientID);

  // Local → Yjs: append every synced mutation to the log. Skipped for
  // view-role clients — their local edits never leave their browser.
  const subDispose = cmdSvc.onMutationExecutedForCollab((info, options) => {
    if (role === 'view') return;
    if (options?.fromCollab) return;
    if (!SYNCED_MUTATIONS.has(info.id)) return;
    const rec: OpRecord = {
      c: myClientId,
      t: Date.now(),
      id: info.id,
      // Univer mutation params are already JSON-friendly (numbers, strings,
      // plain objects). If something carries a Map / Set / cyclic ref we'll
      // discover it via a runtime error; that's the signal to drop the
      // mutation from SYNCED_MUTATIONS.
      p: info.params as unknown,
    };
    // Single-record append — Yjs batches the array operation atomically.
    log.push([rec]);
  });

  // Replay tracking: how many entries we've already executed locally so we
  // don't double-apply on incremental updates. On connect, replay everything
  // we haven't seen — that's how late joiners catch up.
  let appliedCount = 0;

  const replayPending = (): void => {
    const total = log.length;
    while (appliedCount < total) {
      const rec = log.get(appliedCount);
      appliedCount += 1;
      if (!rec) continue;
      if (rec.c === myClientId) continue; // our own write — Univer already ran it
      // Each browser creates its workbook with its OWN random unit id, so
      // raw replay would target the sender's unit (which doesn't exist
      // here) — rewrite to our local active unit. Sheet ids (`sheet-1`)
      // are already deterministic across the room.
      const params = rewriteUnitId(api, rec.p);
      // Fire-and-forget; ordering is preserved by serial awaits not being
      // necessary (each command finishes synchronously enough for the next
      // to start, and Univer's command bus serializes its own dispatch).
      void cmdSvc.executeCommand(rec.id, params, { fromCollab: true });
    }
  };

  const observer = (event: Y.YArrayEvent<OpRecord>) => {
    void event;
    replayPending();
  };
  log.observe(observer);

  // Cover the initial-state case: when the bridge mounts after Yjs has
  // already synced the existing log (provider was connected before us),
  // observe() won't fire — we'd miss everything. Replay synchronously
  // once on mount to catch up.
  replayPending();

  return {
    doc,
    dispose: () => {
      subDispose.dispose();
      log.unobserve(observer);
    },
  };
}

/**
 * Substitute the active local workbook's unit id into a mutation's
 * `params.unitId` so cross-peer mutations target our local workbook.
 * Sheet-level `subUnitId` keys (`sheet-1`, …) are deterministic from
 * the emptyWorkbook snapshot, so they pass through unchanged.
 *
 * Returns a shallow-cloned params object so we don't mutate the Yjs
 * record (Y.Array entries are frozen plain objects).
 */
function rewriteUnitId(api: FUniver, params: unknown): unknown {
  if (!params || typeof params !== 'object') return params;
  const wb = api.getActiveWorkbook();
  if (!wb) return params;
  const localUnitId = wb.getId();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const p = params as Record<string, any>;
  if (typeof p.unitId !== 'string' || p.unitId === localUnitId) return params;
  return { ...p, unitId: localUnitId };
}
