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
 * Signing controller — pure state-machine tests. Mirrors the
 * document/ repo's controller.test.ts; the controller code itself
 * is byte-identical across both products (signing is a uniform
 * concern). When you change controller.ts here, change it there
 * too.
 */
import { strict as assert } from 'node:assert';
import { test } from 'node:test';

import { createSigningController } from './controller';
import type { SignatureField, SignedFieldPayload } from './types';

const ALICE: SignatureField = {
  fieldId: 'alice',
  label: 'Alice',
  required: true,
  anchor: { kind: 'sheet', sheet: 'Sheet1', cell: 'B2' },
  methods: ['drawn'],
};

const BOB: SignatureField = {
  fieldId: 'bob',
  label: 'Bob',
  required: true,
  anchor: { kind: 'sheet', sheet: 'Sheet1', cell: 'B3' },
  methods: ['typed'],
};

const CAROL_OPTIONAL: SignatureField = {
  fieldId: 'carol',
  label: 'Carol (witness)',
  required: false,
  anchor: { kind: 'sheet', sheet: 'Sheet1', cell: 'B4' },
  methods: ['drawn', 'typed'],
};

function payload(fieldId: string): SignedFieldPayload {
  return {
    fieldId,
    method: 'drawn',
    bytes: new Uint8Array([1, 2, 3]).buffer,
    mime: 'image/png',
    signedAt: '2026-06-08T00:00:00Z',
  };
}

test('createSigningController: starts with the first required field active', () => {
  const c = createSigningController([ALICE, BOB], 'sequential');
  assert.equal(c.snapshot().activeFieldIndex, 0);
  assert.equal(c.snapshot().canComplete, false);
  assert.equal(c.snapshot().isComplete, false);
});

test('signField advances activeFieldIndex to the next required', () => {
  const c = createSigningController([ALICE, BOB], 'sequential');
  c.signField(payload('alice'));
  assert.equal(c.snapshot().activeFieldIndex, 1);
  assert.equal(c.snapshot().canComplete, false);
});

test('canComplete flips true once every required field is signed', () => {
  const c = createSigningController([ALICE, BOB, CAROL_OPTIONAL], 'sequential');
  c.signField(payload('alice'));
  c.signField(payload('bob'));
  assert.equal(c.snapshot().canComplete, true);
  assert.equal(c.snapshot().signed['carol'], undefined);
});

test('complete throws when required fields are still unsigned', () => {
  const c = createSigningController([ALICE, BOB], 'sequential');
  c.signField(payload('alice'));
  assert.throws(() => c.complete());
});

test('complete succeeds when all required done', () => {
  const c = createSigningController([ALICE, BOB], 'sequential');
  c.signField(payload('alice'));
  c.signField(payload('bob'));
  const snap = c.complete();
  assert.equal(snap.isComplete, true);
  assert.equal(snap.activeFieldIndex, -1);
});

test('cancel is idempotent and prevents further signField', () => {
  const c = createSigningController([ALICE, BOB], 'sequential');
  c.cancel();
  c.cancel();
  assert.equal(c.snapshot().isCancelled, true);
  c.signField(payload('alice'));
  assert.equal(c.snapshot().signed['alice'], undefined);
});

test('focusField in sequential mode rejects jumps to non-active fields', () => {
  const c = createSigningController([ALICE, BOB], 'sequential');
  c.focusField('bob');
  assert.equal(c.snapshot().activeFieldIndex, 0);
});

test('focusField in concurrent mode allows any unsigned field', () => {
  const c = createSigningController([ALICE, BOB], 'concurrent');
  c.focusField('bob');
  assert.equal(c.snapshot().activeFieldIndex, 1);
});

test('subscribers receive snapshots on every change', () => {
  const c = createSigningController([ALICE, BOB], 'sequential');
  const events: number[] = [];
  const unsub = c.subscribe((s) => events.push(s.activeFieldIndex));
  c.signField(payload('alice'));
  c.signField(payload('bob'));
  c.complete();
  unsub();
  assert.deepEqual(events, [1, -1, -1]);
});

test('signField throws on unknown fieldId', () => {
  const c = createSigningController([ALICE], 'sequential');
  assert.throws(() =>
    c.signField({
      ...payload('not-a-field'),
      fieldId: 'not-a-field',
    }),
  );
});

test('rejects duplicate fieldId at construction', () => {
  assert.throws(() => createSigningController([ALICE, { ...BOB, fieldId: 'alice' }], 'sequential'));
});

test('rejects empty field array at construction', () => {
  assert.throws(() => createSigningController([], 'sequential'));
});
