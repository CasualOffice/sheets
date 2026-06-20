import type * as Y from 'yjs';
import type { FUniver } from '@univerjs/core/facade';
import type { IWorkbookData } from '@univerjs/core';
import { ICommandService, type ICommandInfo, type IExecutionOptions } from '@univerjs/core';
import { SetRangeValuesUndoMutationFactory } from '@univerjs/sheets';
import { deepRewriteUnitId, rewriteJson1OpPathUnitId } from './bridge-helpers';
import { ensurePluginByName, type LazyPluginGroup } from '../univer';
import {
  classifyReplayError,
  pushDeadLetter,
  TRANSIENT_RETRY_DELAYS_MS,
  withRetry,
  type ReplayFailureRecord,
} from './replay-retry';

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
  'sheet.mutation.add-hyper-link': 'hyperlink',
  'sheet.mutation.remove-hyper-link': 'hyperlink',
  'sheet.mutation.update-hyper-link': 'hyperlink',
  'data-validation.mutation.addRule': 'dv',
  'data-validation.mutation.removeRule': 'dv',
  'data-validation.mutation.updateRule': 'dv',
  'sheet.mutation.set-drawing-apply': 'drawing',
};
// y-protocols ships type declarations only as ESM and our tsconfig
// doesn't pick them up cleanly; loose-type the Awareness surface we
// actually use (getStates → Map keyed by clientID).
// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Awareness = { getStates(): Map<number, any> };

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
export const SYNCED_MUTATIONS: ReadonlySet<string> = new Set([
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
  // Sheet visibility — hide/show. NB: `set-worksheet-activate` is
  // deliberately omitted so each peer keeps their own active sheet
  // independent of which sheet another user is editing.
  'sheet.mutation.set-worksheet-hidden',
  // Freeze.
  'sheet.mutation.set-frozen',
  // Hyperlinks (sheets-hyper-link).
  'sheet.mutation.add-hyper-link',
  'sheet.mutation.remove-hyper-link',
  'sheet.mutation.update-hyper-link',
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
  // Drawings / images (sheets-drawing). Single all-purpose mutation
  // wraps add / remove / update via a JSON-1 op + an enum type. Params
  // can be large (embedded image blobs) — accept the bandwidth hit
  // until we move drawings to a side-channel resource model.
  'sheet.mutation.set-drawing-apply',
  // Workbook / worksheet metadata. Each is rarely changed mid-session
  // but cheap to propagate when it does happen. Without these,
  // renaming the workbook or toggling gridlines silently stays
  // local-only — confusing in a shared room.
  // NOTE: `set-worksheet-right-to-left` is intentionally NOT here —
  // neither the command nor the mutation is registered in
  // @univerjs/sheets@0.22.1 (it's exported but never wired up by any
  // plugin), so nothing in our app can emit it. Add it back if a
  // future Univer bump registers it.
  'sheet.mutation.set-workbook-name',
  'sheet.mutation.set-worksheet-row-count',
  'sheet.mutation.set-worksheet-column-count',
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

type MutationRecord = {
  kind?: 'op';
  /** Yjs client id of the emitter (string for portability via JSON). */
  c: string;
  /** Wall-clock at emit; diagnostic only. */
  t: number;
  /** Mutation id (e.g. `sheet.mutation.set-range-values`). */
  id: string;
  /** Mutation params, JSON-serializable. */
  p: unknown;
  /** Optional undo params — set for mutations in REVERTABLE_MUTATIONS,
   *  computed via Univer's `*UndoMutationFactory` BEFORE the redo
   *  runs (so it reads pre-edit state). The HistoryPanel's Revert
   *  button feeds this back into `executeCommand(rec.id, rec.u)` to
   *  restore the pre-edit values. Older log entries without `u`
   *  predate this feature — their Revert button stays disabled. */
  u?: unknown;
};

/**
 * Snapshot entry written into the op log by the designated compactor
 * client (lowest awareness clientId). Replaces all prior entries; any
 * mutation records that come AFTER it in the array are post-compaction
 * incremental edits and replay normally.
 *
 * Pipeline Stage 6 — keeps long-lived rooms from accumulating an
 * unbounded op log. A 24-hour room with light editing could otherwise
 * grow to thousands of records, slowing every late join. Compaction
 * collapses it back to "snapshot + a handful of recent ops".
 */
type SnapshotRecord = {
  kind: 'snapshot';
  c: string;
  t: number;
  /** Full IWorkbookData. Yes, this is large for big workbooks — but
   *  it ships once per compaction interval, not per mutation. The
   *  trade-off vs. unbounded op-log growth is straightforward. */
  wb: IWorkbookData;
};

type OpRecord = MutationRecord | SnapshotRecord;

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
   * designated compactor (lowest known clientId — deterministic and
   * race-free) and (b) so view-only clients don't try to compact. If
   * omitted, compaction is disabled.
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
};

/**
 * Compaction thresholds. We only attempt to compact when the log has
 * grown past `COMPACT_OPS_THRESHOLD` AND at least
 * `COMPACT_MIN_INTERVAL_MS` has elapsed since the last compaction.
 * The interval guard prevents two designated-writer candidates from
 * racing the compaction; the ops threshold avoids compacting a quiet
 * room over and over.
 */
const COMPACT_OPS_THRESHOLD = 200;
const COMPACT_MIN_INTERVAL_MS = 60_000;
const COMPACT_CHECK_INTERVAL_MS = 30_000;

export function startBridge(api: FUniver, doc: Y.Doc, opts: BridgeOptions = {}): BridgeHandle {
  const role = opts.role ?? 'write';
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

  const log = doc.getArray<OpRecord>(LOG_KEY);
  const myClientId = String(doc.clientID);
  // One-shot guard for the __splitChunk__ regression watchdog below.
  let splitChunkWarned = false;

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
    if (options?.fromCollab) return;
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

  // Local → Yjs: append every synced mutation to the log. Skipped for
  // view-role clients — their local edits never leave their browser.
  //
  // We BATCH the appends across a microtask window so a paste / sort
  // that emits many mutations doesn't trigger one Yjs encode per
  // mutation — that path was the main contributor to "large action
  // takes 3–5 s" on big workbooks. Single Y.Array.push with N entries
  // is one transaction, one encode, one WS frame.
  let pending: OpRecord[] = [];
  let flushScheduled = false;
  const flush = () => {
    flushScheduled = false;
    if (pending.length === 0) return;
    const batch = pending;
    pending = [];
    // doc.transact wraps the push in a single transaction so subscribers
    // see one change event for the whole batch.
    doc.transact(() => {
      log.push(batch);
    });
  };
  const subDispose = cmdSvc.onMutationExecutedForCollab((info, options) => {
    if (role === 'view') return;
    if (options?.fromCollab) return;
    if (!SYNCED_MUTATIONS.has(info.id)) return;
    // Univer 0.22.x doesn't use the chunked-mutation protocol that
    // earlier versions (and the CLAUDE.md hard rule) referenced. If a
    // future upgrade reintroduces __splitChunk__, mutations split
    // across multiple emissions will silently corrupt on peers because
    // our op-log doesn't reassemble them. Warn loudly the FIRST time
    // we see one so an upgrade regression surfaces in the console
    // instead of as a mysterious paste-corruption bug.
    if (!splitChunkWarned && hasSplitChunkMarker(info.params)) {
      splitChunkWarned = true;
      console.warn(
        '[collab] mutation "%s" carries __splitChunk__ — Univer reintroduced chunked mutations; bridge needs reassembly logic. See docs/COLLAB-FIXES.md issue 8.',
        info.id,
      );
    }
    // Pair with the undo params we captured at beforeCommandExecuted
    // (only set for REVERTABLE_MUTATIONS). The before-hook ran a few
    // microseconds ago with the SAME params object — match by
    // stringified key.
    const key = JSON.stringify(info.params);
    const undoParams = pendingUndo.get(key);
    pendingUndo.delete(key);
    pending.push({
      c: myClientId,
      t: Date.now(),
      id: info.id,
      // Univer mutation params are already JSON-friendly (numbers, strings,
      // plain objects). If something carries a Map / Set / cyclic ref we'll
      // discover it via a runtime error; that's the signal to drop the
      // mutation from SYNCED_MUTATIONS.
      p: info.params as unknown,
      ...(undoParams !== undefined ? { u: undoParams } : {}),
    });
    if (!flushScheduled) {
      flushScheduled = true;
      // queueMicrotask runs after the current command completes but
      // before the browser paints — keeps the bridge low-latency while
      // letting a multi-mutation command (paste, sort, fill) coalesce.
      queueMicrotask(flush);
    }
  });

  // Replay tracking: how many entries we've already executed locally so we
  // don't double-apply on incremental updates. On connect, replay everything
  // we haven't seen — that's how late joiners catch up.
  let appliedCount = 0;
  // Single-flight guard: when a snapshot record needs an async workbook
  // swap, subsequent records have to wait for the swap to land — otherwise
  // they execute against the old unit id and silently fork state. We also
  // want a single observer callback at a time so re-entrant Yjs events
  // don't interleave half-applied loops.
  let replayInFlight: Promise<void> | null = null;

  const replayPending = (): Promise<void> => {
    if (replayInFlight) return replayInFlight;
    // CRITICAL: assign `replayInFlight = p` BEFORE invoking the async
    // IIFE. The previous version was:
    //   replayInFlight = (async () => { try { ... } finally { replayInFlight = null; } })();
    // For an empty log the IIFE body has no `await` and runs
    // synchronously — the `finally` set `replayInFlight = null`
    // BEFORE the outer `replayInFlight = ...promise...` assignment,
    // which then OVERWROTE the null with the freshly-resolved promise.
    // Result: `replayInFlight` stayed truthy forever and every
    // subsequent `replayPending()` returned immediately without
    // doing anything — remote mutations sat in the Yjs log untouched.
    // Tracker: docs/COLLAB-FIXES.md issue #29.
    let resolveOuter!: () => void;
    const p = new Promise<void>((r) => {
      resolveOuter = r;
    });
    replayInFlight = p;
    void (async () => {
      try {
        // Loop until we catch up. `log.length` may grow while we're awaiting
        // a snapshot apply, so re-read on each pass.

        while (true) {
          const total = log.length;
          // Stage 6 compaction shrinks the log atomically. If our cursor
          // is past the new end, reset to 0 and replay the snapshot record
          // (which is always at position 0 right after compaction).
          if (appliedCount > total) appliedCount = 0;
          if (appliedCount >= total) {
            break;
          }
          const rec = log.get(appliedCount);
          appliedCount += 1;
          if (!rec) continue;
          if (rec.c === myClientId) continue; // our own write — Univer already ran it
          if (rec.kind === 'snapshot') {
            // Compaction record from a peer — replace the local workbook
            // with the snapshot. Without `onSnapshotReceived` wired (e.g.
            // in unit tests that drive the bridge directly), skip the
            // record; the next post-snapshot mutations may still apply
            // cleanly if state is close enough.
            //
            // CRITICAL: await the handler. Univer's unit swap is async, and
            // continuing the loop before the new unit is wired into the
            // facade means rewriteUnitId() reads the OLD active unit and
            // every subsequent mutation targets a stale workbook.
            if (opts.onSnapshotReceived) {
              try {
                await opts.onSnapshotReceived(rec.wb);
              } catch (err) {
                console.warn('[collab] failed to apply compaction snapshot', err);
              }
            } else {
              console.warn(
                '[collab] received compaction snapshot but no handler — workbook may diverge',
              );
            }
            continue;
          }
          // Each browser creates its workbook with its OWN random unit id, so
          // raw replay would target the sender's unit (which doesn't exist
          // here) — rewrite to our local active unit. Sheet ids (`sheet-1`)
          // are already deterministic across the room.
          const params = rewriteUnitId(api, rec.p, rec.id);
          // Univer's ActiveWorksheetController unconditionally switches
          // the active sheet on every insert-sheet mutation — there's no
          // `fromCollab` opt-out inside Univer. Save our current active
          // sheet around the replay and restore it after the next tick
          // so peers don't get yanked to whichever sheet someone else
          // just created.
          const sheetBefore =
            rec.id === 'sheet.mutation.insert-sheet' ? captureActiveSheetId(api) : null;
          // Lazy-plugin gate: if this mutation belongs to a plugin we
          // haven't mounted yet (CF, tables, filter, notes,
          // hyperlinks), the mutation handler is missing and the
          // change drops silently. AWAIT plugin load before executing.
          // For mutations not in the map, this resolves to undefined
          // and the executeCommand fires immediately.
          const lazyGroup = MUTATION_TO_LAZY_GROUP[rec.id];
          // Fire-and-forget; ordering is preserved by Univer's command bus
          // serialising its own dispatch.
          //
          // Failure handling is two-class (see replay-retry.ts):
          //   - TRANSIENT (dynamic-import chunk-load failures) → retry
          //     with 300/900/2700 ms backoff. The lazy-plugin gate is
          //     the common source: a network flap during webpack chunk
          //     fetch rejects the import; retries land cleanly once
          //     connectivity recovers.
          //   - PERMANENT (malformed params, unknown command id, range
          //     out-of-bounds) → dead-letter immediately. Retrying
          //     just re-throws the same stack.
          //
          // Final failure (after retries exhausted OR permanent on
          // first throw) increments `replayFailures` AND appends to
          // the dead-letter ring buffer for the UI to render.
          const attempt = () =>
            (lazyGroup ? ensurePluginByName(lazyGroup) : Promise.resolve()).then(() =>
              cmdSvc.executeCommand(rec.id, params, { fromCollab: true }),
            );
          void withRetry(
            attempt,
            TRANSIENT_RETRY_DELAYS_MS,
            (err) => classifyReplayError(err) === 'transient',
          )
            .then(() => {
              if (sheetBefore) restoreActiveSheetId(api, sheetBefore);
            })
            .catch((err: unknown) => {
              const cls = classifyReplayError(err);
              const message = err instanceof Error ? err.message : String(err);
              console.warn(
                '[collab] replay failed for',
                rec.id,
                '(class:',
                cls + ',',
                'gave up)',
                err,
              );
              const now = Date.now();
              const failure: ReplayFailureRecord = {
                id: rec.id,
                params: rec.p,
                lastError: message,
                // Permanent = 1 attempt; transient = 1 + N retries
                // configured in TRANSIENT_RETRY_DELAYS_MS.
                attempts: cls === 'transient' ? 1 + TRANSIENT_RETRY_DELAYS_MS.length : 1,
                firstFailedAt: now,
                lastFailedAt: now,
                classification: cls,
              };
              noteReplayFailure(failure);
              if (sheetBefore) restoreActiveSheetId(api, sheetBefore);
            });
        }
      } finally {
        // Only clear if we're still the in-flight token. A future
        // re-entrant guard scheme might let multiple flights coexist;
        // this check keeps us correct under that.
        if (replayInFlight === p) replayInFlight = null;
        resolveOuter();
      }
    })();
    return p;
  };

  const observer = (event: Y.YArrayEvent<OpRecord>) => {
    void event;
    void replayPending();
  };
  log.observe(observer);

  // Cover the initial-state case: when the bridge mounts after Yjs has
  // already synced the existing log (provider was connected before us),
  // observe() won't fire — we'd miss everything. Replay synchronously
  // once on mount to catch up.
  void replayPending();

  // ── Stage 6: periodic compaction by the designated writer ────────
  // Only one client in the room compacts at a time — the one with the
  // lowest known clientId. The interval guard prevents an over-eager
  // compactor from churning. View-only clients never compact.
  //
  // Seed `lastCompactedAt` to `now` so the first auto-compaction
  // observes the full COMPACT_MIN_INTERVAL_MS cooldown. Without this
  // seed, a fresh session that immediately crosses the op threshold
  // (e.g. a quick paste of >200 cells, or the e2e harness) would see
  // an instant first compaction before the test could observe the
  // pre-compaction log. The explicit `__bridgeForceCompact` path
  // bypasses this guard, so the test still works.
  let lastCompactedAt = Date.now();
  // Both scheduling paths declared in outer scope so the dispose
  // closure below can clean up either one.
  let intervalHandle: ReturnType<typeof setInterval> | null = null;
  let idleHandle: number | null = null;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const cic = (globalThis as any).cancelIdleCallback as undefined | ((id: number) => void);
  if (role !== 'view' && opts.awareness) {
    const awareness = opts.awareness;
    const tryCompact = (): void => {
      try {
        if (log.length < COMPACT_OPS_THRESHOLD) return;
        if (Date.now() - lastCompactedAt < COMPACT_MIN_INTERVAL_MS) return;
        // Designated writer = lowest clientId currently in awareness.
        // Math.min over the awareness keys, then compare to ours.
        const keys = Array.from(awareness.getStates().keys()) as number[];
        if (keys.length === 0) return;
        const designated = Math.min(...keys);
        if (designated !== doc.clientID) return;
        // Snapshot the live workbook.
        const wb = api.getActiveWorkbook();
        if (!wb) return;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const snap = (wb as any).save() as IWorkbookData;
        const snapshotRec: SnapshotRecord = {
          kind: 'snapshot',
          c: myClientId,
          t: Date.now(),
          wb: snap,
        };
        const opsBefore = log.length;
        // Atomic swap: clear then append. Yjs serializes the whole
        // transaction so subscribers see one consistent change.
        doc.transact(() => {
          log.delete(0, log.length);
          log.push([snapshotRec]);
        });
        // We just rewrote the log; our cursor must stay PAST the
        // snapshot (we already have its state). The replayer's
        // appliedCount > length reset would otherwise re-apply our
        // own snapshot which is a no-op but pointless.
        appliedCount = 1;
        lastCompactedAt = Date.now();
        console.info('[collab] op-log compacted: %d ops → 1 snapshot record', opsBefore);
      } catch (err) {
        console.warn('[collab] compaction attempt failed', err);
      }
    };
    // Schedule `tryCompact` via requestIdleCallback so the heavy
    // `wb.save()` only runs when the main thread is genuinely idle —
    // never mid-keystroke or mid-paste. The browser gives us a
    // deadline; if it expires before we'd start, we skip and wait
    // for the next tick. Fall back to a plain setInterval in
    // environments without rIC (Safari < 18, some test runners).
    //
    // `wb.save()` itself can't move to a Web Worker (the Univer
    // workbook is a main-thread object graph), so the realistic
    // optimisation is "don't run it when the user is busy".
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ric = (globalThis as any).requestIdleCallback as
      | undefined
      | ((
          cb: (d: { didTimeout: boolean; timeRemaining: () => number }) => void,
          opts?: { timeout: number },
        ) => number);
    if (typeof ric === 'function') {
      const scheduleNext = () => {
        idleHandle = ric(
          (deadline) => {
            // Only run if we have at least ~5 ms to spare (typical
            // empty-workbook save is <1 ms, big ones a few ms;
            // anything longer should defer to the next idle window).
            if (deadline.didTimeout || deadline.timeRemaining() > 5) {
              tryCompact();
            }
            scheduleNext();
          },
          { timeout: COMPACT_CHECK_INTERVAL_MS },
        );
      };
      scheduleNext();
    } else {
      intervalHandle = setInterval(tryCompact, COMPACT_CHECK_INTERVAL_MS);
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
      // Bypass the COMPACT_MIN_INTERVAL_MS guard so the test
      // doesn't have to sleep a minute.
      lastCompactedAt = 0;
      tryCompact();
    };
  }

  return {
    doc,
    dispose: () => {
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
  };
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
 * Cheap probe for the `__splitChunk__` marker — a flag Univer used in
 * earlier versions to indicate a mutation was one chunk of a larger
 * operation (large paste, copy-worksheet). Univer 0.22.x doesn't emit
 * it, but if a future upgrade reintroduces it our op-log replay would
 * silently corrupt because we don't reassemble chunks. Watchdog logs
 * a warning the first time it sees one so the regression is loud.
 *
 * Walks one level deep — Univer carried the marker on the top-level
 * params object historically. Deeper nesting would be a different
 * shape and warrants a different fix.
 */
function hasSplitChunkMarker(params: unknown): boolean {
  if (!params || typeof params !== 'object') return false;

  return Object.prototype.hasOwnProperty.call(params, '__splitChunk__');
}
