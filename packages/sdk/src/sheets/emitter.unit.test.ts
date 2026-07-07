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
 * Unit coverage for the typed emitter that backs the unified `.on()/.off()`
 * surface (doc 38 §3). The emitter lives in its own Univer-free module so it is
 * importable in the plain node test runner (importing `./api` would pull the
 * Univer runtime, which the other `*.unit.test.ts` files also avoid).
 */

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createEmitter } from './emitter';

interface Events extends Record<string, (...args: never[]) => void> {
  change: (value: number) => void;
  dirtyChange: (dirty: boolean) => void;
}

test('on dispatches to subscribers and returns an unsubscribe fn', () => {
  const e = createEmitter<Events>();
  const seen: number[] = [];
  const off = e.on('change', (v) => seen.push(v));
  e.emit('change', 1);
  e.emit('change', 2);
  off();
  e.emit('change', 3);
  assert.deepEqual(seen, [1, 2]);
});

test('off removes a specific handler', () => {
  const e = createEmitter<Events>();
  const seen: number[] = [];
  const handler = (v: number) => seen.push(v);
  e.on('change', handler);
  e.off('change', handler);
  e.emit('change', 1);
  assert.deepEqual(seen, []);
});

test('events are isolated per name', () => {
  const e = createEmitter<Events>();
  const changes: number[] = [];
  const dirties: boolean[] = [];
  e.on('change', (v) => changes.push(v));
  e.on('dirtyChange', (d) => dirties.push(d));
  e.emit('change', 7);
  e.emit('dirtyChange', true);
  assert.deepEqual(changes, [7]);
  assert.deepEqual(dirties, [true]);
});

test('emit is a no-op with no subscribers', () => {
  const e = createEmitter<Events>();
  assert.doesNotThrow(() => e.emit('change', 1));
  assert.equal(e.listenerCount('change'), 0);
});

test('listenerCount tracks live subscriptions', () => {
  const e = createEmitter<Events>();
  assert.equal(e.listenerCount('change'), 0);
  const off1 = e.on('change', () => {});
  const off2 = e.on('change', () => {});
  assert.equal(e.listenerCount('change'), 2);
  off1();
  assert.equal(e.listenerCount('change'), 1);
  off2();
  assert.equal(e.listenerCount('change'), 0);
});

test('a throwing handler is isolated — peers still fire', () => {
  const e = createEmitter<Events>();
  const seen: number[] = [];
  const originalError = console.error;
  console.error = () => {}; // silence the isolated-throw log for this test
  try {
    e.on('change', () => {
      throw new Error('boom');
    });
    e.on('change', (v) => seen.push(v));
    assert.doesNotThrow(() => e.emit('change', 9));
  } finally {
    console.error = originalError;
  }
  assert.deepEqual(seen, [9]);
});

test('a handler that unsubscribes mid-dispatch does not skip its peers', () => {
  const e = createEmitter<Events>();
  const seen: string[] = [];
  // 'a' removes itself while the emit is in flight.
  const offA = e.on('change', () => {
    seen.push('a');
    offA();
  });
  e.on('change', () => seen.push('b'));
  e.emit('change', 1);
  // Both run for this emit (dispatch is over a snapshot); 'a' is gone next time.
  assert.deepEqual(seen, ['a', 'b']);
  e.emit('change', 2);
  assert.deepEqual(seen, ['a', 'b', 'b']);
});
