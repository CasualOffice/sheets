/**
 * Replay-failure retry classifier + scheduler.
 *
 * Pure helpers used by the Yjs bridge to recover from transient
 * mutation-replay failures (most often dynamic-import chunk-load
 * errors during the lazy-plugin gate) without conflating them with
 * permanent failures (malformed mutation params, unknown command id).
 *
 * Decoupled from bridge.ts so we can unit-test the classification +
 * backoff logic without standing up a full Yjs doc + Univer engine.
 *
 * See docs/PRODUCTION_PIPELINE.md → Stream A1 for design context.
 */

/** Maximum retry attempts for transient failures. */
export const TRANSIENT_RETRY_DELAYS_MS = [300, 900, 2700] as const;

/** Maximum entries kept in the dead-letter ring buffer. */
export const DEAD_LETTER_CAP = 20;

export type ReplayClassification = 'transient' | 'permanent';

export interface ReplayFailureRecord {
  /** Mutation id (e.g. 'sheet.mutation.set-range-values'). */
  id: string;
  /** Mutation params at time of failure (unfiltered; may be large). */
  params: unknown;
  /** Last error's message (string-coerced for log/UI safety). */
  lastError: string;
  /** Total attempts made (1 = first try, no retries). */
  attempts: number;
  /** ms-since-epoch of the FIRST failure for this record. */
  firstFailedAt: number;
  /** ms-since-epoch of the most recent failure for this record. */
  lastFailedAt: number;
  /** Whether the classifier called this transient or permanent. */
  classification: ReplayClassification;
}

/**
 * Decide whether a replay error is worth retrying.
 *
 * Transient class: dynamic-import / chunk-load failures. These come
 * from the lazy-plugin gate fetching a webpack chunk over the
 * network; a flap on the user's connection drops the fetch and the
 * lazy `import()` rejects with one of several distinct shapes:
 *
 *   - Webpack 5: `ChunkLoadError` with name === 'ChunkLoadError'
 *   - Webpack 4 / vite: 'Loading chunk N failed.'
 *   - Vite / native ESM: 'Failed to fetch dynamically imported module'
 *   - Native fetch under offline: 'NetworkError when attempting to fetch'
 *
 * All four resolve on retry once the network recovers. Everything
 * else (executeCommand rejections from bad params, missing handlers,
 * out-of-bounds ranges) is permanent — retrying just burns the same
 * stack trace N more times.
 *
 * Conservative bias: when in doubt, classify permanent. A false
 * negative (transient → permanent) costs us a dead-letter entry the
 * user can re-trigger by reloading; a false positive (permanent →
 * transient) wastes 4 s of retries on a known-broken mutation.
 */
export function classifyReplayError(err: unknown): ReplayClassification {
  if (err == null) return 'permanent';

  // Error instances: check name first (ChunkLoadError sets this even
  // though it's not on stock Error), then message.
  const e = err as { name?: unknown; message?: unknown };
  const name = typeof e.name === 'string' ? e.name : '';
  const message = typeof e.message === 'string' ? e.message : String(err);

  if (name === 'ChunkLoadError') return 'transient';

  const lower = message.toLowerCase();
  if (lower.includes('loading chunk') && lower.includes('failed')) {
    return 'transient';
  }
  if (lower.includes('failed to fetch dynamically imported')) {
    return 'transient';
  }
  if (lower.includes('networkerror when attempting to fetch')) {
    return 'transient';
  }
  // Generic 'network request failed' from some bundler error wrappers.
  if (lower.includes('network request failed')) return 'transient';

  return 'permanent';
}

/**
 * Run `task` with up to `delays.length` retries. The first attempt is
 * immediate; each subsequent attempt waits `delays[i-1]` ms.
 *
 * Resolves with the task's value on the first success. Rejects with
 * the LAST error after all attempts fail.
 *
 * `shouldRetry` is called between attempts. If it returns false, the
 * loop bails immediately with the current error — this lets the
 * caller switch to permanent-failure handling mid-sequence (e.g. if
 * a follow-up error reveals the real problem was permanent all along).
 */
export async function withRetry<T>(
  task: () => Promise<T>,
  delays: readonly number[],
  shouldRetry: (err: unknown, attemptsSoFar: number) => boolean = () => true,
  sleep: (ms: number) => Promise<void> = (ms) => new Promise((r) => setTimeout(r, ms)),
): Promise<T> {
  let lastErr: unknown;
  for (let i = 0; i <= delays.length; i += 1) {
    try {
      return await task();
    } catch (err) {
      lastErr = err;
      const attempts = i + 1;
      if (i >= delays.length) break;
      if (!shouldRetry(err, attempts)) break;
      await sleep(delays[i]);
    }
  }
  throw lastErr;
}

/**
 * Ring buffer for dead-letter records. Append-only; oldest entry
 * evicts on overflow. Caller owns the backing array — we return a
 * new array on push so React state updates can see the change by
 * reference.
 */
export function pushDeadLetter(
  buffer: readonly ReplayFailureRecord[],
  rec: ReplayFailureRecord,
  cap: number = DEAD_LETTER_CAP,
): ReplayFailureRecord[] {
  const next =
    buffer.length >= cap ? [...buffer.slice(buffer.length - cap + 1), rec] : [...buffer, rec];
  return next;
}
