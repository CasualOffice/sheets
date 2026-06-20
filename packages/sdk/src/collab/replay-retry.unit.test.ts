/**
 * Pure-state tests for the replay-retry classifier + scheduler. The
 * bridge owns the integration (which depends on a live Yjs doc +
 * Univer engine); these tests pin the failure-class boundary so the
 * "what counts as transient" decision is checked-in instead of buried
 * in production code.
 *
 * Run with: `pnpm test:unit`
 */
import { strict as assert } from 'node:assert';
import { test } from 'node:test';

import {
  classifyReplayError,
  DEAD_LETTER_CAP,
  pushDeadLetter,
  TRANSIENT_RETRY_DELAYS_MS,
  withRetry,
  type ReplayFailureRecord,
} from './replay-retry';

/* ── classifyReplayError ──────────────────────────────────────────── */

test('ChunkLoadError by name is transient', () => {
  const err = Object.assign(new Error('Loading chunk 7 failed.'), {
    name: 'ChunkLoadError',
  });
  assert.equal(classifyReplayError(err), 'transient');
});

test('webpack "Loading chunk N failed" message is transient', () => {
  // Bare Error with no special name — only the message tells us.
  assert.equal(classifyReplayError(new Error('Loading chunk 42 failed.')), 'transient');
});

test('vite "failed to fetch dynamically imported module" is transient', () => {
  assert.equal(
    classifyReplayError(new Error('Failed to fetch dynamically imported module: /assets/cf.js')),
    'transient',
  );
});

test('NetworkError when attempting to fetch is transient', () => {
  assert.equal(
    classifyReplayError(new Error('NetworkError when attempting to fetch resource.')),
    'transient',
  );
});

test('generic "network request failed" is transient', () => {
  assert.equal(classifyReplayError(new Error('Network request failed')), 'transient');
});

test('malformed-params TypeError is permanent', () => {
  // Realistic shape: Univer's set-range-values throws TypeError when
  // params.value is undefined.
  assert.equal(
    classifyReplayError(new TypeError("Cannot read properties of undefined (reading 'v')")),
    'permanent',
  );
});

test('unknown command id Error is permanent', () => {
  assert.equal(
    classifyReplayError(new Error('No command handler registered for "made.up.mutation"')),
    'permanent',
  );
});

test('null / undefined / number errors classify permanent (defensive)', () => {
  assert.equal(classifyReplayError(null), 'permanent');
  assert.equal(classifyReplayError(undefined), 'permanent');
  assert.equal(classifyReplayError(42), 'permanent');
});

test('case-insensitive: "LOADING CHUNK 1 FAILED" is transient', () => {
  // The classifier lowercases before matching so we don't miss a
  // bundler that screams its errors. Realistic case for some
  // production-mode bundle wrappers.
  assert.equal(classifyReplayError(new Error('LOADING CHUNK 1 FAILED.')), 'transient');
});

/* ── withRetry ──────────────────────────────────────────────────── */

test('withRetry resolves on first try when task succeeds', async () => {
  let calls = 0;
  const value = await withRetry(
    () => {
      calls += 1;
      return Promise.resolve('ok');
    },
    [10, 20],
    () => true,
    () => Promise.resolve(),
  );
  assert.equal(value, 'ok');
  assert.equal(calls, 1);
});

test('withRetry retries until a later attempt succeeds', async () => {
  let calls = 0;
  const value = await withRetry(
    () => {
      calls += 1;
      if (calls < 3) return Promise.reject(new Error('chunk load failed'));
      return Promise.resolve('eventually');
    },
    [10, 20, 30],
    () => true,
    () => Promise.resolve(),
  );
  assert.equal(value, 'eventually');
  assert.equal(calls, 3);
});

test('withRetry exhausts all attempts then rejects with last error', async () => {
  let calls = 0;
  await assert.rejects(
    withRetry(
      () => {
        calls += 1;
        return Promise.reject(new Error(`attempt ${calls}`));
      },
      [10, 20, 30],
      () => true,
      () => Promise.resolve(),
    ),
    { message: 'attempt 4' }, // 1 initial + 3 retries = 4 calls
  );
  assert.equal(calls, 4);
});

test('withRetry bails early when shouldRetry returns false', async () => {
  let calls = 0;
  await assert.rejects(
    withRetry(
      () => {
        calls += 1;
        return Promise.reject(new Error('permanent'));
      },
      [10, 20, 30],
      () => false, // never retry
      () => Promise.resolve(),
    ),
    { message: 'permanent' },
  );
  assert.equal(calls, 1);
});

test('withRetry actually awaits the sleep delays (in order)', async () => {
  const slept: number[] = [];
  let calls = 0;
  await assert.rejects(
    withRetry(
      () => {
        calls += 1;
        return Promise.reject(new Error('fail'));
      },
      [50, 150, 450],
      () => true,
      (ms) => {
        slept.push(ms);
        return Promise.resolve();
      },
    ),
  );
  assert.deepEqual(slept, [50, 150, 450]);
  assert.equal(calls, 4);
});

/* ── pushDeadLetter ─────────────────────────────────────────────── */

function makeRec(id: string): ReplayFailureRecord {
  return {
    id,
    params: { foo: 1 },
    lastError: 'boom',
    attempts: 1,
    firstFailedAt: 0,
    lastFailedAt: 0,
    classification: 'permanent',
  };
}

test('pushDeadLetter appends below cap', () => {
  const buf = pushDeadLetter([], makeRec('a'));
  assert.equal(buf.length, 1);
  assert.equal(buf[0].id, 'a');
});

test('pushDeadLetter evicts oldest when at cap', () => {
  const small = 3;
  let buf: readonly ReplayFailureRecord[] = [];
  for (const id of ['a', 'b', 'c', 'd', 'e']) {
    buf = pushDeadLetter(buf, makeRec(id), small);
  }
  assert.equal(buf.length, 3);
  assert.deepEqual(
    buf.map((r) => r.id),
    ['c', 'd', 'e'],
  );
});

test('pushDeadLetter returns a new array (reference change)', () => {
  // The CollabIndicator subscribes by reference — we MUST return a
  // fresh array, not mutate in place.
  const original: readonly ReplayFailureRecord[] = [makeRec('a')];
  const next = pushDeadLetter(original, makeRec('b'));
  assert.notEqual(next, original);
  assert.equal(original.length, 1);
  assert.equal(next.length, 2);
});

test('TRANSIENT_RETRY_DELAYS_MS is the documented schedule', () => {
  // Pin the schedule. If this changes, audit the indicator copy +
  // CO-EDITING.md so the user-visible "should recover within ~4 s"
  // claim stays accurate.
  assert.deepEqual([...TRANSIENT_RETRY_DELAYS_MS], [300, 900, 2700]);
});

test('DEAD_LETTER_CAP is 20', () => {
  assert.equal(DEAD_LETTER_CAP, 20);
});
