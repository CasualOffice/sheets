import * as Y from 'yjs';
import type { FUniver } from '@univerjs/core/facade';
import type { IWorkbookData } from '@univerjs/core';
import {
  ICommandService,
  type ICommandInfo,
  type IExecutionOptions,
} from '@univerjs/core';
import { deepRewriteUnitId } from './bridge-helpers';
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
    pending.push({
      c: myClientId,
      t: Date.now(),
      id: info.id,
      // Univer mutation params are already JSON-friendly (numbers, strings,
      // plain objects). If something carries a Map / Set / cyclic ref we'll
      // discover it via a runtime error; that's the signal to drop the
      // mutation from SYNCED_MUTATIONS.
      p: info.params as unknown,
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
    const p = new Promise<void>((r) => { resolveOuter = r; });
    replayInFlight = p;
    void (async () => {
      try {
        // Loop until we catch up. `log.length` may grow while we're awaiting
        // a snapshot apply, so re-read on each pass.
        // eslint-disable-next-line no-constant-condition
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
              console.warn('[collab] received compaction snapshot but no handler — workbook may diverge');
            }
            continue;
          }
          // Each browser creates its workbook with its OWN random unit id, so
          // raw replay would target the sender's unit (which doesn't exist
          // here) — rewrite to our local active unit. Sheet ids (`sheet-1`)
          // are already deterministic across the room.
          const params = rewriteUnitId(api, rec.p);
          // Univer's ActiveWorksheetController unconditionally switches
          // the active sheet on every insert-sheet mutation — there's no
          // `fromCollab` opt-out inside Univer. Save our current active
          // sheet around the replay and restore it after the next tick
          // so peers don't get yanked to whichever sheet someone else
          // just created.
          const sheetBefore =
            rec.id === 'sheet.mutation.insert-sheet' ? captureActiveSheetId(api) : null;
          // Fire-and-forget; ordering is preserved by Univer's command bus
          // serialising its own dispatch.
          // `.catch` surfaces malformed-mutation / out-of-bounds replays
          // that would otherwise silently fork local state — every
          // divergence we've chased after a release started as one of
          // these swallowed throws.
          cmdSvc.executeCommand(rec.id, params, { fromCollab: true })
            .then(() => {
              if (sheetBefore) restoreActiveSheetId(api, sheetBefore);
            })
            .catch((err) => {
              console.warn('[collab] replay failed for', rec.id, err);
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
  let lastCompactedAt = 0;
  let compactionTimer: ReturnType<typeof setInterval> | null = null;
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
        console.info(
          '[collab] op-log compacted: %d ops → 1 snapshot record',
          opsBefore,
        );
      } catch (err) {
        console.warn('[collab] compaction attempt failed', err);
      }
    };
    compactionTimer = setInterval(tryCompact, COMPACT_CHECK_INTERVAL_MS);
    // Don't keep the interval alive in tests / SSR where this
    // module might be imported but never torn down.
    compactionTimer.unref?.();

    // Dev-only sinks for the compaction e2e — lets a test trigger
    // the compaction without waiting for the 30 s interval and read
    // the live log length. Tree-shaken from production builds via
    // import.meta.env.DEV.
    if (import.meta.env.DEV) {
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
  }

  return {
    doc,
    dispose: () => {
      subDispose.dispose();
      // Flush any pending batch so an edit-then-leave race doesn't drop
      // the last keystroke on the floor.
      if (pending.length > 0) flush();
      log.unobserve(observer);
      if (compactionTimer) clearInterval(compactionTimer);
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
function rewriteUnitId(api: FUniver, params: unknown): unknown {
  const wb = api.getActiveWorkbook();
  if (!wb) return params;
  const localUnitId = wb.getId();
  return deepRewriteUnitId(params, localUnitId);
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
