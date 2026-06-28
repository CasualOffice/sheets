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

import { strict as assert } from 'node:assert';
import { afterEach, beforeEach, test } from 'node:test';

import { isMacroMutation, nextMacroName, type Macro } from './macros.js';

test('isMacroMutation: only sheet.mutation.* records', () => {
  assert.equal(isMacroMutation('sheet.mutation.set-range-values'), true);
  assert.equal(isMacroMutation('sheet.mutation.set-range-styles'), true);
  // Transient noise that must NOT be replayed:
  assert.equal(isMacroMutation('formula.mutation.set-formula-calculation-start'), false);
  assert.equal(isMacroMutation('doc.mutation.rich-text-editing'), false);
  assert.equal(isMacroMutation('sheet.command.set-range-values'), false); // command, not mutation
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  assert.equal(isMacroMutation(undefined as any), false);
});

test('nextMacroName increments past existing names', () => {
  const existing: Macro[] = [
    { name: 'Macro 1', steps: [], createdAt: 0 },
    { name: 'Macro 2', steps: [], createdAt: 0 },
  ];
  assert.equal(nextMacroName(existing), 'Macro 3');
  assert.equal(nextMacroName([]), 'Macro 1');
  // Fills the lowest free slot.
  assert.equal(nextMacroName([{ name: 'Macro 2', steps: [], createdAt: 0 }]), 'Macro 1');
});

// localStorage-backed save/list/delete round-trip with a minimal shim.
let store: Record<string, string>;
beforeEach(() => {
  store = {};
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (globalThis as any).localStorage = {
    getItem: (k: string) => (k in store ? store[k] : null),
    setItem: (k: string, v: string) => {
      store[k] = v;
    },
    removeItem: (k: string) => {
      delete store[k];
    },
  };
});
afterEach(() => {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  delete (globalThis as any).localStorage;
});

test('saveMacro / listMacros / deleteMacro round-trip', async () => {
  const { saveMacro, listMacros, deleteMacro } = await import('./macros.js');
  const m: Macro = {
    name: 'Macro 1',
    steps: [{ id: 'sheet.mutation.set-range-values', params: { x: 1 } }],
    createdAt: 123,
  };
  saveMacro(m);
  assert.deepEqual(listMacros(), [m]);
  // Save by same name replaces, not duplicates.
  const m2: Macro = { ...m, createdAt: 456 };
  saveMacro(m2);
  assert.equal(listMacros().length, 1);
  assert.equal(listMacros()[0].createdAt, 456);
  // Delete.
  deleteMacro('Macro 1');
  assert.deepEqual(listMacros(), []);
});

test('renameMacro: preserves order/steps/shortcut, rejects blank + duplicate', async () => {
  const { saveMacro, listMacros, renameMacro, setMacroShortcut } = await import('./macros.js');
  saveMacro({
    name: 'Macro 1',
    steps: [{ id: 'sheet.mutation.set-range-values', params: { x: 1 } }],
    createdAt: 10,
  });
  saveMacro({ name: 'Macro 2', steps: [], createdAt: 20 });
  setMacroShortcut('Macro 1', 'M');

  // Happy path: rename keeps position, steps, shortcut, createdAt.
  renameMacro('Macro 1', 'Monthly close');
  const after = listMacros();
  assert.deepEqual(
    after.map((m) => m.name),
    ['Monthly close', 'Macro 2'],
    'order preserved (rename in place, not delete+append)',
  );
  const renamed = after.find((m) => m.name === 'Monthly close')!;
  assert.equal(renamed.shortcut, 'M');
  assert.equal(renamed.createdAt, 10);
  assert.equal(renamed.steps.length, 1);

  // Trims whitespace.
  renameMacro('Monthly close', '  Weekly  ');
  assert.ok(listMacros().some((m) => m.name === 'Weekly'));

  // Rejections — list unchanged.
  const names = () => listMacros().map((m) => m.name);
  renameMacro('Weekly', '   '); // blank
  assert.deepEqual(names(), ['Weekly', 'Macro 2']);
  renameMacro('Weekly', 'Macro 2'); // duplicate
  assert.deepEqual(names(), ['Weekly', 'Macro 2']);
  renameMacro('Nope', 'Whatever'); // old name not found
  assert.deepEqual(names(), ['Weekly', 'Macro 2']);
});

test('setMacroShortcut assigns, is unique, and rejects reserved letters', async () => {
  const { saveMacro, setMacroShortcut, findMacroByShortcut, listMacros } =
    await import('./macros.js');
  saveMacro({ name: 'A', steps: [], createdAt: 0 });
  saveMacro({ name: 'B', steps: [], createdAt: 0 });

  // Assign + lookup (case-insensitive input, stored uppercase).
  setMacroShortcut('A', 'm');
  assert.equal(findMacroByShortcut('M')?.name, 'A');
  assert.equal(listMacros().find((m) => m.name === 'A')?.shortcut, 'M');

  // Uniqueness: assigning M to B steals it from A.
  setMacroShortcut('B', 'M');
  assert.equal(findMacroByShortcut('M')?.name, 'B');
  assert.equal(listMacros().find((m) => m.name === 'A')?.shortcut, undefined);

  // Reserved letters (L/D/P) are rejected — binding unchanged.
  setMacroShortcut('A', 'L');
  assert.equal(findMacroByShortcut('L'), undefined);
  assert.equal(listMacros().find((m) => m.name === 'A')?.shortcut, undefined);

  // Clear with null.
  setMacroShortcut('B', null);
  assert.equal(findMacroByShortcut('M'), undefined);
});

test('availableMacroLetters excludes reserved + other-macro letters, keeps own', async () => {
  const { saveMacro, setMacroShortcut, availableMacroLetters } = await import('./macros.js');
  saveMacro({ name: 'A', steps: [], createdAt: 0 });
  saveMacro({ name: 'B', steps: [], createdAt: 0 });
  setMacroShortcut('A', 'M');

  const forB = availableMacroLetters('B');
  assert.ok(!forB.includes('L') && !forB.includes('D') && !forB.includes('P')); // reserved
  assert.ok(!forB.includes('M')); // held by A
  assert.ok(forB.includes('Z'));

  // A's own list still includes M (so its current binding shows as selected).
  assert.ok(availableMacroLetters('A').includes('M'));
});
