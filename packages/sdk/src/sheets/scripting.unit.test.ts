/**
 * Unit tests for the pure scripting primitives behind CasualSheetsAPI's
 * `executeCommands` (batch replay) and `onMutation` (record/observe). These
 * are the primitives the host scripting API — and the app's macro feature —
 * are built on.
 *
 * Imports only `./scripting` (no `@univerjs/*` value imports), so it runs under
 * the bare `node --import tsx` runner. The thin facade wiring in `api.ts` is
 * covered by the SDK e2e harness.
 *
 * Run with: `pnpm test:unit`
 */
import { strict as assert } from 'node:assert';
import { test } from 'node:test';

import {
  attachMutationObserver,
  runSteps,
  type CommandRecord,
  type MutationEmitter,
} from './scripting';

test('runSteps replays every step in order and counts them', async () => {
  const calls: CommandRecord[] = [];
  const steps: CommandRecord[] = [
    { id: 'sheet.mutation.set-range-values', params: { a: 1 } },
    { id: 'sheet.mutation.set-range-values', params: { a: 2 } },
    { id: 'sheet.mutation.insert-row', params: { r: 3 } },
  ];
  const applied = await runSteps((id, params) => {
    calls.push({ id, params });
    return true;
  }, steps);
  assert.equal(applied, 3);
  assert.deepEqual(calls, steps);
});

test('runSteps is best-effort: a throwing step is skipped, not fatal', async () => {
  const calls: string[] = [];
  const applied = await runSteps(
    (id) => {
      if (id === 'sheet.mutation.boom') throw new Error('boom');
      calls.push(id);
      return true;
    },
    [
      { id: 'sheet.mutation.set-range-values', params: { a: 1 } },
      { id: 'sheet.mutation.boom' },
      { id: 'sheet.mutation.insert-row', params: { r: 3 } },
    ],
  );
  assert.equal(applied, 2);
  assert.deepEqual(calls, ['sheet.mutation.set-range-values', 'sheet.mutation.insert-row']);
});

test('runSteps awaits async execute and on an empty list resolves to 0', async () => {
  assert.equal(await runSteps(async () => true, []), 0);
  let ran = 0;
  const applied = await runSteps(async () => {
    await Promise.resolve();
    ran += 1;
  }, [{ id: 'a' }, { id: 'b' }]);
  assert.equal(applied, 2);
  assert.equal(ran, 2);
});

/** A fake command service whose collab hook we can fire manually. */
function fakeCmdSvc() {
  let listener: ((info: CommandRecord) => void) | undefined;
  let disposed = false;
  const svc: MutationEmitter = {
    onMutationExecutedForCollab(l) {
      listener = l;
      return {
        dispose() {
          listener = undefined;
          disposed = true;
        },
      };
    },
  };
  return {
    svc,
    emit: (info: CommandRecord) => listener?.(info),
    isObserving: () => listener !== undefined,
    wasDisposed: () => disposed,
  };
}

test('attachMutationObserver forwards {id, params} and the disposer stops the stream', () => {
  const { svc, emit, isObserving, wasDisposed } = fakeCmdSvc();
  const seen: CommandRecord[] = [];
  const stop = attachMutationObserver(svc, (r) => seen.push(r));
  assert.equal(isObserving(), true);

  emit({ id: 'sheet.mutation.set-range-values', params: { v: 5 } });
  emit({ id: 'sheet.mutation.insert-row', params: { r: 1 } });
  assert.deepEqual(seen, [
    { id: 'sheet.mutation.set-range-values', params: { v: 5 } },
    { id: 'sheet.mutation.insert-row', params: { r: 1 } },
  ]);

  stop();
  assert.equal(wasDisposed(), true);
  emit({ id: 'sheet.mutation.set-range-values', params: { v: 9 } });
  assert.equal(seen.length, 2, 'no events after dispose');
});

test('attachMutationObserver tolerates an absent service (no-op disposer)', () => {
  const stop = attachMutationObserver(undefined, () => {
    throw new Error('should never be called');
  });
  assert.doesNotThrow(stop);
});

test('observe → replay round-trips a recorded mutation stream', async () => {
  const { svc, emit, isObserving } = fakeCmdSvc();
  const recorded: CommandRecord[] = [];
  const stop = attachMutationObserver(svc, (r) => recorded.push(r));
  emit({ id: 'sheet.mutation.set-range-values', params: { a: 1 } });
  emit({ id: 'sheet.mutation.set-range-values', params: { a: 2 } });
  stop();
  assert.equal(isObserving(), false);

  const replayed: CommandRecord[] = [];
  const applied = await runSteps((id, params) => replayed.push({ id, params }), recorded);
  assert.equal(applied, 2);
  assert.deepEqual(replayed, recorded);
});
