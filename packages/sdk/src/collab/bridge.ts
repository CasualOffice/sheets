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

import type * as Y from 'yjs';
import type { FUniver } from '@univerjs/core/facade';
import type { IWorkbookData } from '@univerjs/core';
import { ICommandService, type ICommandInfo, type IExecutionOptions } from '@univerjs/core';
import { SetRangeValuesUndoMutationFactory } from '@univerjs/sheets';
import type { CasualSheetsAPI } from '../sheets/api';
import { deepRewriteUnitId, rewriteJson1OpPathUnitId } from './bridge-helpers';
import { ensurePluginByName, type LazyPluginGroup } from '../univer';
import {
  classifyReplayError,
  pushDeadLetter,
  TRANSIENT_RETRY_DELAYS_MS,
  withRetry,
  type ReplayFailureRecord,
} from './replay-retry';
import {
  buildReplayPlan,
  COLLAB_LOG_PROTOCOL_VERSION,
  CollabProtocolError,
  isPrefix,
  mergeFrontier,
  recordId,
  type MutationRecord,
  type OpRecord,
  type ReplayPlan,
  type SnapshotRecord,
} from './replay-plan';

/**
 * Map mutation ids to the lazy-plugin group that owns the matching
 * mutation handler. The joiner replays peer mutations through Univer's
 * command service; if the receiving plugin hasn't been loaded yet
 * (lazy bundling), the mutation handler is missing and the change
 * silently drops on that peer. Bridge waits for the plugin to mount
 * before executing the mutation.
 */
const MUTATION_TO_LAZY_GROUP: Record<string, LazyPluginGroup> = {
  'sheet.mutation.add-conditional-rule': 'cf',
  'sheet.mutation.set-conditional-rule': 'cf',
  'sheet.mutation.delete-conditional-rule': 'cf',
  'sheet.mutation.move-conditional-rule': 'cf',
  'sheet.mutation.add-table': 'table',
  'sheet.mutation.delete-table': 'table',
  'sheet.mutation.set-sheet-table': 'table',
  'sheet.mutation.set-table-filter': 'table',
  'sheet.mutation.set-filter-criteria': 'filter',
  'sheet.mutation.set-filter-range': 'filter',
  'sheet.mutation.remove-filter': 'filter',
  'sheet.mutation.update-note': 'note',
  'sheet.mutation.remove-note': 'note',
  'sheets.mutation.add-hyper-link': 'hyperlink',
  'sheets.mutation.remove-hyper-link': 'hyperlink',
  'sheets.mutation.update-hyper-link': 'hyperlink',
  'data-validation.mutation.addRule': 'dv',
  'data-validation.mutation.removeRule': 'dv',
  'data-validation.mutation.updateRule': 'dv',
  'sheet.mutation.set-drawing-apply': 'drawing',
  // Thread comments (thread-comment + sheets-thread-comment). The
  // mutation handlers live in @univerjs/thread-comment; the
  // sheets-thread-comment(-ui) plugins are the lazy-loaded integration
  // our `threadComment` group registers. A joiner that hasn't opened
  // the Comments pane yet has none of these mounted, so the replay
  // would drop the peer's comment silently — gate on plugin load.
  'thread-comment.mutation.add-comment': 'threadComment',
  'thread-comment.mutation.update-comment': 'threadComment',
  'thread-comment.mutation.update-comment-ref': 'threadComment',
  'thread-comment.mutation.resolve-comment': 'threadComment',
  'thread-comment.mutation.delete-comment': 'threadComment',
};
// y-protocols ships type declarations only as ESM and our tsconfig
// doesn't pick them up cleanly; loose-type the Awareness surface we
// actually use (getStates → Map keyed by clientID).
// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Awareness = { getStates(): Map<number, any> };

/**
 * Yjs ↔ Univer mutation bridge. See docs/CO-EDITING.md for the design.
 *
 * Strategy: every non-collab mutation gets a stable client/sequence id and is
 * serialized into a Y.Array log. Peers replay the logical v2 plan strictly in
 * order with `fromCollab: true`. If a concurrent insert changes the already
 * applied prefix, the bridge restores a known workbook base and replays the
 * canonical Yjs order so optimistic local execution cannot fork peers.
 *
 * Why an op log (not a state mirror): writing a per-mutation state mirror
 * for every mutation Univer emits (set-range-values, set-style, insert-row,
 * merge, hide-col, freeze, …) is dozens of handlers. The log generalizes
 * — any deterministic mutation just round-trips its params. Trade-off: no
 * per-cell CRDT merging on concurrent writes (Yjs orders inserts, then
 * Univer re-executes them; last writer wins at the mutation level).
 *
 * Echo-loop guard (per CLAUDE.md):
 *   - Records carry a stable id. A local record is skipped only while the
 *     canonical prefix has not moved; after a base restore it is replayed too.
 *   - Remote applies pass `fromCollab: true` so Univer's
 *     `onMutationExecutedForCollab` listener filters them back out via the
 *     private options identity check below.
 *   - Univer mutations marked `onlyLocal` or `fromChangeset` are deliberately
 *     non-replicating runtime/snapshot work and never become source log ops.
 */

const LOG_KEY = 'ops';
const COMPACTION_ORIGIN = Symbol('casual-sheets-compaction-v2');

/**
 * Allowlist of mutation ids we sync. Listed explicitly to keep
 * undocumented / version-volatile mutations out of the log. An unlisted
 * state mutation marks the browser unsafe for compaction so it cannot become
 * a divergent room base. Easier to add new ids than to debug corruption from
 * a mutation that secretly references local state (render skeletons, etc.).
 */
export const SYNCED_MUTATIONS: ReadonlySet<string> = new Set([
  // Cell-level — values, formulas, styles, rich text.
  'sheet.mutation.set-range-values',
  // Number formats are separate eager-plugin mutations in Univer 0.25.
  'sheet.mutation.set.numfmt',
  'sheet.mutation.remove.numfmt',
  // Row / column structural.
  'sheet.mutation.insert-row',
  'sheet.mutation.insert-col',
  'sheet.mutation.remove-rows',
  'sheet.mutation.remove-col',
  'sheet.mutation.move-rows',
  'sheet.mutation.move-columns',
  'sheet.mutation.set-row-hidden',
  'sheet.mutation.set-row-visible',
  'sheet.mutation.set-col-hidden',
  'sheet.mutation.set-col-visible',
  'sheet.mutation.set-worksheet-row-height',
  'sheet.mutation.set-worksheet-row-is-auto-height',
  'sheet.mutation.set-worksheet-row-auto-height',
  'sheet.mutation.set-worksheet-col-width',
  // Merges.
  'sheet.mutation.add-worksheet-merge',
  'sheet.mutation.remove-worksheet-merge',
  // Sheet lifecycle.
  'sheet.mutation.insert-sheet',
  'sheet.mutation.remove-sheet',
  'sheet.mutation.set-worksheet-name',
  'sheet.mutation.set-worksheet-order',
  // Sheet visibility — hide/show. NB: `set-worksheet-activate` is
  // deliberately omitted so each peer keeps their own active sheet
  // independent of which sheet another user is editing.
  'sheet.mutation.set-worksheet-hidden',
  // Freeze.
  'sheet.mutation.set-frozen',
  // Hyperlinks (sheets-hyper-link).
  'sheets.mutation.add-hyper-link',
  'sheets.mutation.remove-hyper-link',
  'sheets.mutation.update-hyper-link',
  // Tab colour — picks up the right-click "Tab color" menu.
  'sheet.mutation.set-tab-color',
  // Move + sort. Without these, a peer's cut-and-paste-cell-block or
  // sort-range action silently doesn't appear on receivers.
  'sheet.mutation.move-range',
  'sheet.mutation.reorder-range',
  // Per-row / per-column metadata (height, custom style, colData).
  // The narrower set-worksheet-row-height / set-worksheet-col-width
  // are already allowlisted; these are the broader resource-style
  // mutations Univer emits for "Format → Row/Column" operations.
  'sheet.mutation.set-row-data',
  'sheet.mutation.set-col-data',
  'sheet.mutation.set-worksheet-default-style',
  // Format-as-table / sheets-table — adds/removes named tables and
  // their config. Picked up so the table chrome appears on both peers.
  'sheet.mutation.add-table',
  'sheet.mutation.delete-table',
  'sheet.mutation.set-sheet-table',
  'sheet.mutation.set-table-filter',
  // Autofilter (sheets-filter).
  'sheet.mutation.set-filter-criteria',
  'sheet.mutation.set-filter-range',
  'sheet.mutation.remove-filter',
  // Notes (sheets-note) — the small cell-corner indicator + popup.
  'sheet.mutation.update-note',
  'sheet.mutation.remove-note',
  // Conditional formatting (sheets-conditional-formatting). Mutations
  // are self-contained — `add` carries the full rule, `set` carries
  // the patched rule, `delete` / `move` carry rule ids. Univer's
  // `ConditionalFormattingRuleModel` consumes them and triggers a
  // canvas re-render so highlighted cells update on peers. Existing
  // rules in a downloaded seed already load via the workbook's
  // resource channel; the mutations cover deltas during the session.
  'sheet.mutation.add-conditional-rule',
  'sheet.mutation.set-conditional-rule',
  'sheet.mutation.delete-conditional-rule',
  'sheet.mutation.move-conditional-rule',
  // Data validation (data-validation core). NB: this package uses the
  // `data-validation.mutation.*` prefix, not `sheet.mutation.*`.
  // Mutation handlers live in @univerjs/data-validation; the
  // sheets-data-validation plugin is the lazy-loaded integration our
  // MUTATION_TO_LAZY_GROUP map keys on.
  'data-validation.mutation.addRule',
  'data-validation.mutation.removeRule',
  'data-validation.mutation.updateRule',
  // Thread comments (thread-comment + sheets-thread-comment). The
  // five mutations are self-contained and JSON-friendly: `add-comment`
  // carries the full IThreadComment (client-generated id + threadId, so
  // concurrent adds converge), `update-comment` / `resolve-comment` /
  // `delete-comment` carry the comment id, and `update-comment-ref`
  // carries the new A1 ref when an edit moves the anchored cell. All
  // use `unitId` (rewritten to the local unit) + `subUnitId` (the
  // deterministic sheet id), so they replay cleanly on peers. Without
  // these, a comment added / resolved / deleted in a shared room never
  // reaches co-editors — the in-cell thread and the Comments pane stay
  // out of sync. Existing comments in a downloaded seed already load
  // via the workbook's thread-comment resource channel; these cover
  // the deltas during the session.
  'thread-comment.mutation.add-comment',
  'thread-comment.mutation.update-comment',
  'thread-comment.mutation.update-comment-ref',
  'thread-comment.mutation.resolve-comment',
  'thread-comment.mutation.delete-comment',
  // Drawings / images (sheets-drawing). Single all-purpose mutation
  // wraps add / remove / update via a JSON-1 op + an enum type. Params
  // can be large (embedded image blobs) — accept the bandwidth hit
  // until we move drawings to a side-channel resource model.
  'sheet.mutation.set-drawing-apply',
  // Workbook / worksheet metadata. Each is rarely changed mid-session
  // but cheap to propagate when it does happen. Without these,
  // renaming the workbook or toggling gridlines silently stays
  // local-only — confusing in a shared room.
  'sheet.mutation.set-workbook-name',
  'sheet.mutation.set-worksheet-row-count',
  'sheet.mutation.set-worksheet-column-count',
  'sheet.mutation.set-worksheet-right-to-left',
  'sheet.mutation.toggle-gridlines',
  'sheet.mutation.set-gridlines-color',
]);

/** Mutation ids for which the bridge captures undo params before the
 *  redo runs. Used by the HistoryPanel's revert action. Restricted to
 *  cell-level mutations because the existing Univer factories cover
 *  them and they're the dominant case for "undo my edit"; structural
 *  ops (insert-row / move-range / sort) need their own factories and
 *  are out of scope for v1 revert.
 */
export const REVERTABLE_MUTATIONS: ReadonlySet<string> = new Set([
  'sheet.mutation.set-range-values',
]);

export type BridgeHandle = {
  /** Underlying Yjs document — exposed so tests / devtools can introspect. */
  doc: Y.Doc;
  /** Stop listening and detach. */
  dispose: () => void;
  /**
   * How many remote mutations have thrown during replay since the
   * bridge started. Each one is a candidate divergence — the local
   * state didn't accept the peer's change, so the two state vectors
   * are now off by at least that mutation. The CollabDriver
   * subscribes (see `subscribeReplayFailures`) so the indicator can
   * warn the user before they discover it the hard way.
   */
  getReplayFailures: () => number;
  /**
   * Subscribe to replay-failure count changes. Fires after every
   * increment with the new total. Returns a teardown that unhooks
   * the subscriber. No initial-value fire — caller can read
   * `getReplayFailures()` once at subscribe time if they need it.
   */
  subscribeReplayFailures: (cb: (count: number) => void) => () => void;
  /**
   * Snapshot of the dead-letter ring buffer — mutations that
   * exhausted retries (transient class) or failed immediately
   * (permanent class). Capped at DEAD_LETTER_CAP entries; oldest
   * evicts on overflow. UI consumes this to render the per-failure
   * detail panel.
   */
  getReplayDeadLetter: () => readonly ReplayFailureRecord[];
  /**
   * Subscribe to dead-letter changes. Fires after every push with a
   * fresh array (reference change — React state updates see it).
   * Returns a teardown.
   */
  subscribeReplayDeadLetter: (cb: (entries: readonly ReplayFailureRecord[]) => void) => () => void;
  /** Resolves after the currently-known log is applied or replay is blocked. */
  whenReplaySettled: () => Promise<void>;
  /** Whether replay is halted at an unapplied protocol/mutation failure. */
  isReplayBlocked: () => boolean;
  /** Diagnostic/manual compaction hook. Returns false when safety gates fail. */
  forceCompact: () => boolean;
};

export type BridgeOptions = {
  /**
   * `view` clients only RECEIVE remote updates — local mutations don't
   * append to the log, so peers never see them. Client-side gate only;
   * a determined user can run the bridge in write mode by editing the
   * URL. Server-side enforcement is a follow-up hardening pass.
   */
  role?: 'view' | 'write';
  /**
   * Provider's Yjs awareness. Used (a) to determine which peer is the
   * designated compactor inside one connected awareness view (lowest known
   * clientId) and (b) so view-only clients don't try to compact. Awareness is
   * not partition-safe election. If omitted, compaction is disabled.
   */
  awareness?: Awareness;
  /**
   * Hand a fresh workbook snapshot to the host when a compaction
   * record arrives from a peer. The bridge can't call
   * `replaceWorkbook` directly (it lives in React state); the host
   * (`CollabDriver`) wires this through.
   *
   * MAY return a promise. Replay of subsequent op-log entries is
   * paused until the promise resolves — without that, mutations
   * land on the OLD unit before Univer's async unit-swap completes,
   * which silently forks state on late joiners.
   */
  onSnapshotReceived?: (wb: IWorkbookData) => void | Promise<void>;
  /**
   * Preservation-aware snapshot reader. Required for safe reorder recovery and
   * compaction when `api` is a bare FUniver facade. Passing CasualSheetsAPI
   * supplies `api.getContent()` automatically.
   */
  getContent?: () => IWorkbookData | null;
  /**
   * Client compaction policy. `off` keeps replay/awareness but prevents every
   * browser snapshot write (use this when the host owns validated checkpoints).
   * `manual` exposes only `forceCompact`; `auto` opts into the idle timer.
   * Defaults to `off`: awareness cannot elect one writer safely across a
   * network partition, so automatic browser checkpoints must be explicit.
   */
  compaction?: 'auto' | 'manual' | 'off';
};

/**
 * Compaction thresholds. We only attempt to compact when the log has
 * grown past `COMPACT_OPS_THRESHOLD` AND at least
 * `COMPACT_MIN_INTERVAL_MS` has elapsed since the last compaction.
 * The interval guard limits checkpoint churn after awareness changes; it is
 * not partition-safe consensus (which is why auto mode is opt-in). The ops
 * threshold avoids compacting a quiet room over and over.
 */
const COMPACT_OPS_THRESHOLD = 200;
const COMPACT_MIN_INTERVAL_MS = 60_000;
const COMPACT_CHECK_INTERVAL_MS = 30_000;

export type BridgeAttachable = FUniver | CasualSheetsAPI;

export function startBridge(
  attachable: BridgeAttachable,
  doc: Y.Doc,
  opts: BridgeOptions = {},
): BridgeHandle {
  const api = resolveFacade(attachable);
  const role = opts.role ?? 'write';
  const compactionMode = opts.compaction ?? 'off';
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const injector = (api as any)._injector as { get: (token: unknown) => unknown } | undefined;
  if (!injector) {
    throw new Error('[collab] FUniver injector not accessible — Univer too old?');
  }
  const cmdSvc = injector.get(ICommandService) as {
    onMutationExecutedForCollab: (l: (info: ICommandInfo, options?: IExecutionOptions) => void) => {
      dispose: () => void;
    };
    beforeCommandExecuted: (l: (info: ICommandInfo, options?: IExecutionOptions) => void) => {
      dispose: () => void;
    };
    executeCommand: (id: string, params: unknown, options?: IExecutionOptions) => Promise<unknown>;
  };
  // `fromCollab` is part of Univer's public options surface, so any host caller
  // can forge it. Use object identity as the bridge-owned capability for echo
  // suppression; keep the boolean only as the downstream Univer signal.
  const replayExecutionOptions: IExecutionOptions = { fromCollab: true };

  const log = doc.getArray<OpRecord>(LOG_KEY);
  const myClientId = String(doc.clientID);
  const casualApi = isCasualSheetsAPI(attachable) ? attachable : null;
  const readContent = opts.getContent ?? (casualApi ? () => casualApi.getContent() : null);
  const applySnapshot =
    opts.onSnapshotReceived ?? (casualApi ? (wb: IWorkbookData) => casualApi.setContent(wb) : null);
  let initialBase: IWorkbookData | null = null;
  if (readContent) {
    try {
      const content = readContent();
      initialBase = content ? cloneWorkbook(content) : null;
    } catch (err) {
      console.warn('[collab] failed to capture the initial replay base', err);
    }
  }
  let nextSequence = 0;
  let disposed = false;

  // Replay-failure tracking — surfaces silent divergences to the UI.
  // Every time a remote mutation throws on apply, this counter ticks
  // and any subscribers (CollabDriver) get the new total. Local writes
  // never tick this (they're applied by Univer before we even append
  // to the log).
  let replayFailures = 0;
  const replayFailureSubscribers = new Set<(count: number) => void>();
  let deadLetter: readonly ReplayFailureRecord[] = [];
  const deadLetterSubscribers = new Set<(entries: readonly ReplayFailureRecord[]) => void>();
  const noteReplayFailure = (rec: ReplayFailureRecord) => {
    replayFailures += 1;
    deadLetter = pushDeadLetter(deadLetter, rec);
    for (const cb of replayFailureSubscribers) {
      try {
        cb(replayFailures);
      } catch (err) {
        console.warn('[collab] replay-failure subscriber threw', err);
      }
    }
    for (const cb of deadLetterSubscribers) {
      try {
        cb(deadLetter);
      } catch (err) {
        console.warn('[collab] dead-letter subscriber threw', err);
      }
    }
  };
  // Undo params keyed by JSON.stringify(params) so we can pair them up
  // when the matching `onMutationExecutedForCollab` fires moments later.
  // Cleared after each pairing — there's no eviction policy because the
  // window between before-execute and after-execute is microseconds.
  const pendingUndo = new Map<string, unknown>();
  // Capture undo params BEFORE the redo runs. The factory walks the
  // current cell state and produces a redo-shaped object that would
  // restore those cells. Only set-range-values for v1; other types fall
  // through with no `u` field — the HistoryPanel disables Revert.
  const subBeforeDispose = cmdSvc.beforeCommandExecuted((info, options) => {
    if (role === 'view') return;
    if (options === replayExecutionOptions || options?.onlyLocal || options?.fromChangeset) return;
    if (!REVERTABLE_MUTATIONS.has(info.id)) return;
    try {
      // The factory's first arg is described as "accessor" — Univer's
      // accessor IS the injector for our purposes (both expose .get).
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const undo = SetRangeValuesUndoMutationFactory(injector as any, info.params as any);
      pendingUndo.set(JSON.stringify(info.params), undo);
    } catch (err) {
      // Pre-edit state was unreadable (workbook missing, sheet gone,
      // etc.) — skip. The history entry just won't be revertable.
      console.warn('[collab] failed to capture undo params for', info.id, err);
    }
  });

  // Stable replay state. Positional cursors are unsafe for a Y.Array: a
  // concurrent append can merge before the cursor and be skipped forever.
  // We remember the exact logical record-id prefix applied to this workbook.
  // If Yjs later inserts before it, restore a known base and replay the final
  // deterministic order.
  let appliedIds: string[] = [];
  const locallyAppliedIds = new Set<string>();
  // A counter, rather than a boolean, matters across an awaited workbook swap:
  // a local edit can request ANOTHER rebuild while the current rebuild is in
  // flight. Clearing a boolean after the await would erase that request and
  // let the optimistic edit be skipped even though the swap discarded it.
  let rebuildGeneration = 0;
  let handledRebuildGeneration = 0;
  let replayBlocked: { key: string; message: string } | null = null;
  let reportedBlockedKey: string | null = null;
  let replayInFlight: Promise<void> | null = null;
  let replayRequested = false;
  let logRevision = 0;
  let localStateUnsafe = false;
  let localUnsafeGeneration = 0;
  let appendBlocked = false;

  const requestRebuild = (): void => {
    rebuildGeneration += 1;
  };
  const needsRebuild = (): boolean => handledRebuildGeneration < rebuildGeneration;

  const reportBlocked = (
    key: string,
    id: string,
    params: unknown,
    err: unknown,
    attempts: number,
  ): void => {
    const message = err instanceof Error ? err.message : String(err);
    replayBlocked = { key, message };
    console.warn('[collab] replay blocked at', id, err);
    if (reportedBlockedKey === key) return;
    reportedBlockedKey = key;
    const now = Date.now();
    noteReplayFailure({
      id,
      params,
      lastError: message,
      attempts,
      firstFailedAt: now,
      lastFailedAt: now,
      classification: classifyReplayError(err),
    });
  };

  const clearBlocked = (key?: string): void => {
    if (!key && (localStateUnsafe || appendBlocked)) return;
    if (key && replayBlocked?.key !== key) return;
    replayBlocked = null;
    reportedBlockedKey = null;
  };

  // Local → Yjs. Appends are batched across one microtask, but every record
  // gets a stable per-client sequence before entering the batch.
  let pending: OpRecord[] = [];
  let flushScheduled = false;
  const flush = () => {
    flushScheduled = false;
    if (pending.length === 0) return;
    const batch = pending;
    pending = [];
    const recoveringAppend = appendBlocked;
    try {
      doc.transact(() => {
        log.push(batch);
      });
      appendBlocked = false;
      clearBlocked('local-log-append');
      // A prior append failure means these optimistic edits may have
      // interleaved with remote work while absent from the canonical log.
      // Request this only AFTER the write succeeds; replayPending's drain flag
      // makes the already-fired Y.Array observer loop back and rebuild.
      if (recoveringAppend) {
        requestRebuild();
        void replayPending();
      }
    } catch (err) {
      // The local workbook already contains these edits. Keep them queued and
      // block compaction instead of silently publishing a snapshot without them.
      pending = [...batch, ...pending];
      appendBlocked = true;
      reportBlocked('local-log-append', '__local_log_append__', batch, err, 1);
    }
  };
  const subDispose = cmdSvc.onMutationExecutedForCollab((info, options) => {
    if (role === 'view') return;
    if (options === replayExecutionOptions || options?.onlyLocal || options?.fromChangeset) return;
    if (!SYNCED_MUTATIONS.has(info.id)) {
      // A state mutation outside the protocol makes this browser an unsafe
      // compaction source. Surface it and keep receiving peer ops, but never
      // publish this local workbook as the room base. Do not put it in the
      // remote-replay dead-letter feed: formula/runtime mutations are expected
      // to remain local, and labelling one as a peer edit that failed is false.
      localStateUnsafe = true;
      localUnsafeGeneration += 1;
      console.warn('[collab] local mutation is not collaboration-safe:', info.id);
      return;
    }
    // Pair with the undo params we captured at beforeCommandExecuted
    // (only set for REVERTABLE_MUTATIONS). The before-hook ran a few
    // microseconds ago with the SAME params object — match by
    // stringified key.
    const key = JSON.stringify(info.params);
    const undoParams = pendingUndo.get(key);
    pendingUndo.delete(key);
    const rec: MutationRecord = {
      v: COLLAB_LOG_PROTOCOL_VERSION,
      kind: 'op',
      c: myClientId,
      s: nextSequence++,
      t: Date.now(),
      id: info.id,
      // Univer mutation params are already JSON-friendly (numbers, strings,
      // plain objects). If something carries a Map / Set / cyclic ref we'll
      // discover it via a runtime error; that's the signal to drop the
      // mutation from SYNCED_MUTATIONS.
      p: info.params as unknown,
      ...(undoParams !== undefined ? { u: undoParams } : {}),
    };
    locallyAppliedIds.add(recordId(rec));
    if (replayInFlight || replayBlocked) requestRebuild();
    if (hasSplitChunkMarker(info.params)) {
      // Univer 0.25 marks each independent slice of a large paste/copy with
      // __splitChunk__ so collaboration transports it as a separate changeset
      // instead of one oversized frame. Preserve that boundary in Yjs while
      // stable ids keep the slices strictly ordered on replay.
      if (pending.length > 0) flush();
      pending.push(rec);
      flush();
      return;
    }
    pending.push(rec);
    if (!flushScheduled) {
      flushScheduled = true;
      // queueMicrotask runs after the current command completes but
      // before the browser paints — keeps the bridge low-latency while
      // letting a multi-mutation command (paste, sort, fill) coalesce.
      queueMicrotask(flush);
    }
  });

  const applyMutation = async (rec: MutationRecord): Promise<boolean> => {
    const stableId = recordId(rec);
    if (!SYNCED_MUTATIONS.has(rec.id)) {
      reportBlocked(
        stableId,
        rec.id,
        rec.p,
        new CollabProtocolError(`unsupported collaboration mutation ${rec.id}`),
        1,
      );
      return false;
    }

    const params = rewriteUnitId(api, rec.p, rec.id);
    const sheetBefore = rec.id === 'sheet.mutation.insert-sheet' ? captureActiveSheetId(api) : null;
    const lazyGroup = MUTATION_TO_LAZY_GROUP[rec.id];
    const attempt = async (): Promise<void> => {
      if (lazyGroup) await ensurePluginByName(lazyGroup);
      const result = await cmdSvc.executeCommand(rec.id, params, replayExecutionOptions);
      if (result === false) throw new Error(`mutation handler rejected ${rec.id}`);
    };
    try {
      await withRetry(
        attempt,
        TRANSIENT_RETRY_DELAYS_MS,
        (err) => classifyReplayError(err) === 'transient',
      );
      clearBlocked(stableId);
      return true;
    } catch (err) {
      // A mutation handler is expected to be atomic, but fail closed even if a
      // future handler mutates partially before throwing/returning false. The
      // next retry starts from the known base and replays the successful prefix.
      requestRebuild();
      const classification = classifyReplayError(err);
      reportBlocked(
        stableId,
        rec.id,
        rec.p,
        err,
        classification === 'transient' ? 1 + TRANSIENT_RETRY_DELAYS_MS.length : 1,
      );
      return false;
    } finally {
      if (sheetBefore) restoreActiveSheetId(api, sheetBefore);
    }
  };

  const applyReplayBase = async (wb: IWorkbookData, stableId: string): Promise<boolean> => {
    if (!applySnapshot) {
      reportBlocked(
        stableId,
        '__snapshot__',
        undefined,
        new CollabProtocolError(
          'replay requires a snapshot applier; pass CasualSheetsAPI or onSnapshotReceived',
        ),
        1,
      );
      return false;
    }
    const unsafeAtStart = localUnsafeGeneration;
    try {
      await applySnapshot(cloneWorkbook(wb));
      // A full base restore discards any unsupported local-only mutation that
      // previously made this browser unsafe as a compaction source. Do not
      // clear one that occurred WHILE the async swap was in flight: it may
      // have landed on the newly mounted workbook and is not in the log.
      if (localUnsafeGeneration === unsafeAtStart) localStateUnsafe = false;
      clearBlocked(stableId);
      return true;
    } catch (err) {
      reportBlocked(stableId, '__snapshot__', undefined, err, 1);
      return false;
    }
  };

  const readPlan = (): ReplayPlan | null => {
    try {
      const plan = buildReplayPlan(log.toArray());
      for (const rec of plan.records) {
        if (rec.c === myClientId) nextSequence = Math.max(nextSequence, rec.s + 1);
      }
      return plan;
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      reportBlocked(`protocol:${message}`, '__protocol__', log.toArray(), err, 1);
      return null;
    }
  };

  const runReplay = async (): Promise<void> => {
    let replayOwnMutations = false;
    let cachedPlan: ReplayPlan | null = null;
    let cachedPlanRevision = -1;
    while (!disposed) {
      // Rebuilding the full validated plan after every awaited command makes a
      // late join O(n²). Reuse it while the Y.Array is unchanged; the observer
      // increments logRevision synchronously, so any concurrent insert still
      // forces a fresh plan before the next record is touched.
      if (!cachedPlan || cachedPlanRevision !== logRevision) {
        cachedPlan = readPlan();
        if (!cachedPlan) return;
        cachedPlanRevision = logRevision;
      }
      const plan = cachedPlan;

      if (needsRebuild() || !isPrefix(appliedIds, plan.ids)) {
        const rebuildTarget = rebuildGeneration;
        // A snapshot in the plan is itself the reset base and is applied below.
        // Otherwise restore the preservation-aware workbook captured on attach.
        if (!plan.snapshot) {
          if (!initialBase) {
            reportBlocked(
              'missing-initial-base',
              '__snapshot__',
              undefined,
              new CollabProtocolError(
                'concurrent replay reorder requires api.getContent() captured at attach time',
              ),
              1,
            );
            return;
          }
          if (!(await applyReplayBase(initialBase, 'initial-base'))) return;
        }
        appliedIds = [];
        replayOwnMutations = true;
        // Preserve any newer request raised while applyReplayBase awaited.
        handledRebuildGeneration = Math.max(handledRebuildGeneration, rebuildTarget);
      }

      if (appliedIds.length === plan.ids.length) {
        clearBlocked();
        return;
      }

      if (plan.snapshot && appliedIds.length === 0) {
        const snapshotId = recordId(plan.snapshot);
        if (!(await applyReplayBase(plan.snapshot.wb, snapshotId))) return;
        appliedIds.push(snapshotId);
        replayOwnMutations = true;
        continue;
      }

      const mutationIndex = appliedIds.length - (plan.snapshot ? 1 : 0);
      const rec = plan.mutations[mutationIndex];
      if (!rec) {
        reportBlocked(
          'invalid-replay-plan',
          '__protocol__',
          plan.ids,
          new CollabProtocolError('replay plan and applied prefix are inconsistent'),
          1,
        );
        return;
      }
      const stableId = recordId(rec);
      if (!replayOwnMutations && locallyAppliedIds.has(stableId)) {
        locallyAppliedIds.delete(stableId);
        appliedIds.push(stableId);
        clearBlocked(stableId);
        continue;
      }
      if (!(await applyMutation(rec))) return;
      locallyAppliedIds.delete(stableId);
      appliedIds.push(stableId);
      // Re-read after every await. An update may have inserted before this
      // record while a plugin or command handler was resolving.
    }
  };

  const replayPending = (): Promise<void> => {
    replayRequested = true;
    if (replayInFlight) return replayInFlight;
    const flight = (async () => {
      do {
        replayRequested = false;
        try {
          await runReplay();
        } catch (err) {
          reportBlocked('unexpected-replay-error', '__replay__', undefined, err, 1);
        }
        // An observer can fire while runReplay is awaiting a plugin, command,
        // or snapshot. Drain that request before publishing a settled flight.
      } while (replayRequested && !disposed);
    })();
    replayInFlight = flight;
    void flight.finally(() => {
      if (replayInFlight !== flight) return;
      replayInFlight = null;
      // Close the narrow microtask race where an observer sees this already-
      // resolved flight after the loop's final condition but before finally.
      if (replayRequested && !disposed) void replayPending();
    });
    return flight;
  };

  const observer = (event: Y.YArrayEvent<OpRecord>) => {
    if (event.transaction.origin === COMPACTION_ORIGIN) return;
    logRevision += 1;
    void replayPending();
  };
  log.observe(observer);

  // Cover the initial-state case: when the bridge mounts after Yjs has
  // already synced the existing log (provider was connected before us),
  // observe() won't fire — we'd miss everything. Replay synchronously
  // once on mount to catch up.
  void replayPending();

  // ── Periodic, preservation-aware compaction ─────────────────────
  // A browser may compact only after it has applied the exact current replay
  // plan. Any pending, failed, reordered, or still-running replay makes the
  // attempt a no-op. This prevents a divergent browser from becoming the next
  // room base.
  let lastCompactedAt = Date.now();
  let intervalHandle: ReturnType<typeof setInterval> | null = null;
  let idleHandle: number | null = null;
  let tryCompact = (_ignoreCooldown = false): boolean => false;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const cic = (globalThis as any).cancelIdleCallback as undefined | ((id: number) => void);
  if (role !== 'view' && opts.awareness && readContent && compactionMode !== 'off') {
    const awareness = opts.awareness;
    tryCompact = (ignoreCooldown = false): boolean => {
      try {
        if (disposed) return false;
        if (log.length < COMPACT_OPS_THRESHOLD) return false;
        if (!ignoreCooldown && Date.now() - lastCompactedAt < COMPACT_MIN_INTERVAL_MS) return false;
        if (
          replayInFlight ||
          replayBlocked ||
          needsRebuild() ||
          localStateUnsafe ||
          appendBlocked
        ) {
          return false;
        }
        if (pending.length > 0 || flushScheduled) return false;
        const keys = Array.from(awareness.getStates().keys()) as number[];
        if (keys.length === 0) return false;
        const designated = Math.min(...keys);
        if (designated !== doc.clientID) return false;

        const plan = readPlan();
        if (!plan || replayBlocked) return false;
        if (
          appliedIds.length !== plan.ids.length ||
          !appliedIds.every((id, index) => plan.ids[index] === id)
        ) {
          return false;
        }

        // Never call raw FWorkbook.save() here. CasualSheetsAPI.getContent()
        // merges its opaque-resource shadow before returning the snapshot.
        const content = readContent();
        if (!content) return false;
        const snap = cloneWorkbook(content);
        const frontier = mergeFrontier(plan.snapshot?.frontier, plan.records);
        nextSequence = Math.max(nextSequence, (frontier[myClientId] ?? -1) + 1);
        const snapshotRec: SnapshotRecord = {
          v: COLLAB_LOG_PROTOCOL_VERSION,
          kind: 'snapshot',
          c: myClientId,
          s: nextSequence++,
          t: Date.now(),
          wb: snap,
          frontier,
        };
        const opsBefore = log.length;
        doc.transact(() => {
          log.delete(0, log.length);
          log.push([snapshotRec]);
        }, COMPACTION_ORIGIN);
        appliedIds = [recordId(snapshotRec)];
        for (const id of locallyAppliedIds) {
          const [clientId, sequence] = splitRecordId(id);
          if (sequence <= (frontier[clientId] ?? -1)) locallyAppliedIds.delete(id);
        }
        clearBlocked();
        lastCompactedAt = Date.now();
        console.info('[collab] op-log compacted: %d ops → 1 snapshot record', opsBefore);
        return true;
      } catch (err) {
        console.warn('[collab] compaction attempt failed', err);
        return false;
      }
    };

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ric = (globalThis as any).requestIdleCallback as
      | undefined
      | ((
          cb: (d: { didTimeout: boolean; timeRemaining: () => number }) => void,
          opts?: { timeout: number },
        ) => number);
    if (compactionMode === 'auto' && typeof ric === 'function') {
      const scheduleNext = () => {
        idleHandle = ric(
          (deadline) => {
            // Only run if we have at least ~5 ms to spare (typical
            // empty-workbook save is <1 ms, big ones a few ms;
            // anything longer should defer to the next idle window).
            if (deadline.didTimeout || deadline.timeRemaining() > 5) {
              tryCompact(false);
            }
            if (!disposed) scheduleNext();
          },
          { timeout: COMPACT_CHECK_INTERVAL_MS },
        );
      };
      scheduleNext();
    } else if (compactionMode === 'auto') {
      intervalHandle = setInterval(() => tryCompact(false), COMPACT_CHECK_INTERVAL_MS);
      intervalHandle.unref?.();
    }

    // Diagnostic sinks for the compaction e2e — lets a test trigger
    // compaction without waiting for the 30 s interval and read the live
    // log length. Installed unconditionally: now that the bridge ships
    // from the SDK bundle (not the app's Vite source), there's no reliable
    // `import.meta.env.DEV` to gate on, and these are the same class of
    // read-only/devtools escape hatch the host already exposes via
    // `__univerAPI` / `__hocuspocusProvider`. No secrets; `__bridgeForceCompact`
    // only triggers the compaction that runs on its own anyway.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (globalThis as any).__bridgeLogLength = () => log.length;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (globalThis as any).__bridgeForceCompact = () => {
      return tryCompact(true);
    };
  } else if (role !== 'view' && opts.awareness && !readContent && compactionMode !== 'off') {
    console.warn(
      '[collab] compaction disabled: pass CasualSheetsAPI or BridgeOptions.getContent so opaque resources are preserved',
    );
  }

  return {
    doc,
    dispose: () => {
      disposed = true;
      subDispose.dispose();
      subBeforeDispose.dispose();
      // Flush any pending batch so an edit-then-leave race doesn't drop
      // the last keystroke on the floor.
      if (pending.length > 0) flush();
      log.unobserve(observer);
      pendingUndo.clear();
      replayFailureSubscribers.clear();
      deadLetterSubscribers.clear();
      // Two scheduling paths in setup above — clean up whichever ran.
      if (intervalHandle) clearInterval(intervalHandle);
      if (idleHandle !== null && typeof cic === 'function') cic(idleHandle);
    },
    getReplayFailures: () => replayFailures,
    subscribeReplayFailures: (cb) => {
      replayFailureSubscribers.add(cb);
      return () => {
        replayFailureSubscribers.delete(cb);
      };
    },
    getReplayDeadLetter: () => deadLetter,
    subscribeReplayDeadLetter: (cb) => {
      deadLetterSubscribers.add(cb);
      return () => {
        deadLetterSubscribers.delete(cb);
      };
    },
    whenReplaySettled: async () => {
      // Calling this also retries a blocked head, which lets a host resume
      // after registering a lazily-loaded command/plugin.
      // Flush synchronously instead of waiting for the queued microtask so the
      // returned promise really covers every local edit known at call time.
      if (pending.length > 0) flush();
      await replayPending();
      while (replayInFlight) await replayInFlight;
    },
    isReplayBlocked: () => replayBlocked !== null || localStateUnsafe || appendBlocked,
    forceCompact: () => tryCompact(true),
  };
}

function isCasualSheetsAPI(api: BridgeAttachable): api is CasualSheetsAPI {
  const candidate = api as Partial<CasualSheetsAPI>;
  return (
    typeof candidate.getContent === 'function' &&
    typeof candidate.setContent === 'function' &&
    !!candidate.univer
  );
}

function resolveFacade(api: BridgeAttachable): FUniver {
  return isCasualSheetsAPI(api) ? api.univer : api;
}

function cloneWorkbook(wb: IWorkbookData): IWorkbookData {
  if (typeof structuredClone === 'function') return structuredClone(wb);
  return JSON.parse(JSON.stringify(wb)) as IWorkbookData;
}

function splitRecordId(id: string): [string, number] {
  const separator = id.lastIndexOf(':');
  return [id.slice(0, separator), Number(id.slice(separator + 1))];
}

/**
 * Substitute the active local workbook's unit id into a mutation's
 * `unitId` fields — top-level AND nested ones (e.g. `range.unitId`,
 * `source.unitId`, `target.unitId`) — so cross-peer mutations target
 * our local workbook. Sheet-level `subUnitId` keys (`sheet-1`, …)
 * are deterministic from the emptyWorkbook snapshot, so they pass
 * through unchanged.
 *
 * Returns a structurally cloned params object so we don't mutate the
 * Yjs record (Y.Array entries are frozen plain objects). Walks objects
 * and arrays recursively; stops at non-plain values (strings, numbers,
 * dates, etc.).
 *
 * Performance: most mutations have shallow params — the recursive walk
 * adds microseconds. Per-cell value maps stay shallow because they're
 * indexed by stringified row/col, not nested objects.
 */
function rewriteUnitId(api: FUniver, params: unknown, mutationId?: string): unknown {
  const wb = api.getActiveWorkbook();
  if (!wb) return params;
  const localUnitId = wb.getId();
  // Capture the sender's unitId BEFORE deepRewriteUnitId swaps it —
  // drawing mutations need it to patch the json1 op path (which
  // carries unitId in position [0] of a positional array, out of
  // deepRewriteUnitId's reach since it only rewrites object KEYS
  // named `unitId`). See bridge-helpers.ts → rewriteJson1OpPathUnitId.
  //
  // Stream F1 fix: without this, set-drawing-apply replays on a
  // joiner with the OWNER's unitId still embedded in the op,
  // json1.type.apply walks a path that doesn't exist locally,
  // throws a bare "Error" with no message, classifier lands it as
  // PERMANENT, and the drawing silently fails to propagate.
  let senderUnitId: string | undefined;
  if (mutationId === 'sheet.mutation.set-drawing-apply' && params && typeof params === 'object') {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const u = (params as any).unitId;
    if (typeof u === 'string') senderUnitId = u;
  }
  const rewritten = deepRewriteUnitId(params, localUnitId) as unknown;
  if (!senderUnitId || senderUnitId === localUnitId) return rewritten;
  // Drawing mutations only — patch the op's positional path[0].
  if (rewritten && typeof rewritten === 'object') {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const r = rewritten as any;
    const fixedOp = rewriteJson1OpPathUnitId(r.op, senderUnitId, localUnitId);
    if (fixedOp !== r.op) {
      return { ...r, op: fixedOp };
    }
  }
  return rewritten;
}

/**
 * Save the local active sheet id before replaying a `fromCollab`
 * mutation that Univer's controllers may use as a side-channel signal
 * to switch sheets (notably `insert-sheet` — see
 * ActiveWorksheetController in @univerjs/sheets). Returning `null`
 * means "couldn't read, don't try to restore".
 */
function captureActiveSheetId(api: FUniver): string | null {
  try {
    const wb = api.getActiveWorkbook();
    if (!wb) return null;
    const sheet = wb.getActiveSheet();
    if (!sheet) return null;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const id = (sheet as any).getSheetId?.() ?? (sheet as any).getId?.() ?? null;
    return typeof id === 'string' ? id : null;
  } catch {
    return null;
  }
}

function restoreActiveSheetId(api: FUniver, sheetId: string): void {
  try {
    const wb = api.getActiveWorkbook();
    if (!wb) return;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const current = (wb.getActiveSheet() as any)?.getSheetId?.();
    if (current === sheetId) return; // nothing to do
    const sheets = wb.getSheets();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const target = sheets.find((s: any) => s.getSheetId?.() === sheetId);
    if (!target) return; // sheet got deleted in the meantime — leave Univer's choice alone
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (wb as any).setActiveSheet?.(target);
  } catch (err) {
    console.warn('[collab] failed to restore active sheet after remote insert-sheet', err);
  }
}

/**
 * Cheap probe for Univer's top-level `__splitChunk__` transport marker.
 * Each marked mutation is already an independently replayable slice; the
 * bridge only needs to keep it in a separate Yjs transaction.
 */
function hasSplitChunkMarker(params: unknown): boolean {
  if (!params || typeof params !== 'object') return false;

  return Object.prototype.hasOwnProperty.call(params, '__splitChunk__');
}
