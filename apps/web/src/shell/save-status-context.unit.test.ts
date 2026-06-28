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

import { describe, it } from 'node:test';
import assert from 'node:assert/strict';

// Pure-logic test of the reducer-ish semantics — exercises the same
// `markDirty` rule a real component relies on without spinning up
// React. The actual hook is exercised by the e2e spec.

import { type SaveStatus } from './save-status-context';

function applyMarkDirty(prev: SaveStatus): SaveStatus {
  // Mirrors the production guard: only knock saved / error back to idle.
  if (prev.kind === 'saved' || prev.kind === 'error') return { kind: 'idle' };
  return prev;
}

describe('save-status-context — markDirty semantics', () => {
  it('idle → idle (no-op)', () => {
    assert.deepEqual(applyMarkDirty({ kind: 'idle' }), { kind: 'idle' });
  });

  it('saving → saving (no-op — let save outcome land first)', () => {
    assert.deepEqual(applyMarkDirty({ kind: 'saving' }), { kind: 'saving' });
  });

  it('saved → idle (the pill must stop lying once the user edits)', () => {
    const next = applyMarkDirty({ kind: 'saved', savedAt: 1700000000000 });
    assert.deepEqual(next, { kind: 'idle' });
  });

  it('error → idle (a fresh edit invalidates the failure surface)', () => {
    const next = applyMarkDirty({ kind: 'error', message: 'network' });
    assert.deepEqual(next, { kind: 'idle' });
  });
});
