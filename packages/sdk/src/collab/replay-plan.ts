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

import type { IWorkbookData } from '@univerjs/core';

/**
 * Version 2 gives every log record a stable per-client clock. Version 1 used
 * only a positional cursor, so a concurrent Y.Array insertion before that
 * cursor could be skipped forever.
 */
export const COLLAB_LOG_PROTOCOL_VERSION = 2 as const;

type RecordClock = {
  v: typeof COLLAB_LOG_PROTOCOL_VERSION;
  /** Yjs client id of the writer, encoded as a string for JSON portability. */
  c: string;
  /** Strictly increasing sequence within `c`. */
  s: number;
  /** Wall-clock timestamp; diagnostic only. */
  t: number;
};

export type MutationRecord = RecordClock & {
  kind: 'op';
  id: string;
  p: unknown;
  u?: unknown;
};

export type SnapshotRecord = RecordClock & {
  kind: 'snapshot';
  wb: IWorkbookData;
  /**
   * Highest record clock represented by `wb`, keyed by client id. An offline
   * operation that arrives after compaction is replayed only when its clock is
   * beyond this frontier, independent of where Yjs positions it in the array.
   */
  frontier: Record<string, number>;
};

export type OpRecord = MutationRecord | SnapshotRecord;

export type ReplayPlan = {
  /** Dominating compacted base, when the room has compacted at least once. */
  snapshot: SnapshotRecord | null;
  /** Mutations not already represented by the selected snapshot. */
  mutations: MutationRecord[];
  /** Stable logical order used to detect inserts before the applied prefix. */
  ids: string[];
  /** Every validated physical record currently in the Y.Array. */
  records: OpRecord[];
};

export class CollabProtocolError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'CollabProtocolError';
  }
}

export function recordId(rec: Pick<RecordClock, 'c' | 's'>): string {
  return `${rec.c}:${rec.s}`;
}

export function isCoveredByFrontier(
  rec: Pick<RecordClock, 'c' | 's'>,
  frontier: Readonly<Record<string, number>>,
): boolean {
  return rec.s <= (frontier[rec.c] ?? -1);
}

/**
 * Turn the physical CRDT array into one deterministic logical replay plan.
 *
 * A compaction snapshot supersedes records at or below its frontier. If two
 * snapshots are concurrent (neither covers the other), there is no sound way
 * to merge their already-materialized workbooks as mutation logs, so the
 * bridge fails closed instead of picking whichever snapshot arrived first.
 */
export function buildReplayPlan(input: readonly unknown[]): ReplayPlan {
  const records = input.map((value, index) => validateRecord(value, index));
  const seen = new Set<string>();
  for (const rec of records) {
    const id = recordId(rec);
    if (seen.has(id)) {
      throw new CollabProtocolError(`duplicate collaboration record id ${id}`);
    }
    seen.add(id);
  }

  const snapshots = records.filter((rec): rec is SnapshotRecord => rec.kind === 'snapshot');
  let snapshot: SnapshotRecord | null = null;
  if (snapshots.length === 1) {
    snapshot = snapshots[0];
  } else if (snapshots.length > 1) {
    const candidates = snapshots.filter((candidate) =>
      snapshots.every(
        (other) =>
          candidate === other ||
          (isCoveredByFrontier(other, candidate.frontier) &&
            frontierCovers(candidate.frontier, other.frontier)),
      ),
    );
    if (candidates.length !== 1) {
      throw new CollabProtocolError(
        'concurrent compaction snapshots cannot be merged safely; start a fresh room',
      );
    }
    snapshot = candidates[0];
  }

  const mutations = records.filter(
    (rec): rec is MutationRecord =>
      rec.kind === 'op' && (!snapshot || !isCoveredByFrontier(rec, snapshot.frontier)),
  );
  const ids = [...(snapshot ? [recordId(snapshot)] : []), ...mutations.map((rec) => recordId(rec))];
  return { snapshot, mutations, ids, records };
}

export function isPrefix(prefix: readonly string[], complete: readonly string[]): boolean {
  if (prefix.length > complete.length) return false;
  return prefix.every((id, index) => complete[index] === id);
}

export function mergeFrontier(
  base: Readonly<Record<string, number>> | undefined,
  records: readonly Pick<RecordClock, 'c' | 's'>[],
): Record<string, number> {
  const merged: Record<string, number> = { ...(base ?? {}) };
  for (const rec of records) {
    merged[rec.c] = Math.max(merged[rec.c] ?? -1, rec.s);
  }
  return merged;
}

function frontierCovers(
  candidate: Readonly<Record<string, number>>,
  other: Readonly<Record<string, number>>,
): boolean {
  return Object.entries(other).every(
    ([clientId, sequence]) => sequence <= (candidate[clientId] ?? -1),
  );
}

function validateRecord(value: unknown, index: number): OpRecord {
  if (!isPlainObject(value)) {
    throw new CollabProtocolError(`record ${index} is not an object`);
  }
  if (value.v !== COLLAB_LOG_PROTOCOL_VERSION) {
    throw new CollabProtocolError(
      `record ${index} uses unsupported collaboration protocol ${String(value.v ?? 1)}; start a fresh room`,
    );
  }
  if (typeof value.c !== 'string' || value.c.length === 0) {
    throw new CollabProtocolError(`record ${index} has no client id`);
  }
  if (!Number.isSafeInteger(value.s) || (value.s as number) < 0) {
    throw new CollabProtocolError(`record ${index} has an invalid sequence`);
  }
  if (typeof value.t !== 'number' || !Number.isFinite(value.t)) {
    throw new CollabProtocolError(`record ${index} has an invalid timestamp`);
  }

  if (value.kind === 'op') {
    if (typeof value.id !== 'string' || value.id.length === 0) {
      throw new CollabProtocolError(`record ${index} has no mutation id`);
    }
    if (!Object.hasOwn(value, 'p')) {
      throw new CollabProtocolError(`record ${index} has no mutation params`);
    }
    return value as MutationRecord;
  }

  if (value.kind === 'snapshot') {
    if (!isPlainObject(value.wb)) {
      throw new CollabProtocolError(`snapshot ${index} has no workbook`);
    }
    if (!isPlainObject(value.frontier)) {
      throw new CollabProtocolError(`snapshot ${index} has no frontier`);
    }
    for (const [clientId, sequence] of Object.entries(value.frontier)) {
      if (
        !clientId ||
        typeof sequence !== 'number' ||
        !Number.isSafeInteger(sequence) ||
        sequence < 0
      ) {
        throw new CollabProtocolError(`snapshot ${index} has an invalid frontier`);
      }
    }
    const snapshot = value as SnapshotRecord;
    if (isCoveredByFrontier(snapshot, snapshot.frontier)) {
      throw new CollabProtocolError(`snapshot ${index} incorrectly covers itself`);
    }
    return snapshot;
  }

  throw new CollabProtocolError(`record ${index} has unknown kind ${String(value.kind)}`);
}

function isPlainObject(value: unknown): value is Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false;
  const proto = Object.getPrototypeOf(value);
  return proto === Object.prototype || proto === null;
}
