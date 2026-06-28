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
import { formatShortcut, isMacPlatform } from './shortcut-format';

describe('isMacPlatform', () => {
  it('returns true for canonical Mac platform strings', () => {
    assert.equal(isMacPlatform('MacIntel'), true);
    assert.equal(isMacPlatform('iPhone'), true);
    assert.equal(isMacPlatform('iPad'), true);
  });
  it('returns false for Windows / Linux strings', () => {
    assert.equal(isMacPlatform('Win32'), false);
    assert.equal(isMacPlatform('Linux x86_64'), false);
  });
});

describe('formatShortcut', () => {
  it('passes plain Ctrl+X through on Win/Linux', () => {
    assert.equal(formatShortcut('Ctrl+X', 'Win32'), 'Ctrl+X');
    assert.equal(formatShortcut('Ctrl+Shift+V', 'Linux x86_64'), 'Ctrl+Shift+V');
  });

  it('collapses Ctrl+X to ⌘X on Mac', () => {
    assert.equal(formatShortcut('Ctrl+X', 'MacIntel'), '⌘X');
    assert.equal(formatShortcut('Ctrl+S', 'MacIntel'), '⌘S');
  });

  it('orders mac modifiers in Apple-HIG order ⌃⌥⇧⌘', () => {
    // Canonical Ctrl+Shift+V → ⇧⌘V on Mac (Shift before Cmd).
    assert.equal(formatShortcut('Ctrl+Shift+V', 'MacIntel'), '⇧⌘V');
    // Ctrl+Alt+V → ⌥⌘V (Alt before Cmd, no Shift).
    assert.equal(formatShortcut('Ctrl+Alt+V', 'MacIntel'), '⌥⌘V');
  });

  it('renders friendly key names (PgUp / Esc / Space)', () => {
    assert.equal(formatShortcut('Ctrl+PageUp', 'Win32'), 'Ctrl+PgUp');
    assert.equal(formatShortcut('Ctrl+PageDown', 'MacIntel'), '⌘PgDn');
    assert.equal(formatShortcut('Escape', 'Win32'), 'Esc');
    assert.equal(formatShortcut('Ctrl+Space', 'Win32'), 'Ctrl+Space');
  });

  it('passes single F-keys through on both platforms', () => {
    assert.equal(formatShortcut('F2', 'MacIntel'), 'F2');
    assert.equal(formatShortcut('F2', 'Win32'), 'F2');
    assert.equal(formatShortcut('Shift+F11', 'MacIntel'), '⇧F11');
  });

  it('renders the literal + key from a trailing ++ (Ctrl++ = insert cells)', () => {
    assert.equal(formatShortcut('Ctrl++', 'Win32'), 'Ctrl++');
    assert.equal(formatShortcut('Ctrl++', 'MacIntel'), '⌘+');
    // The minus counterpart was never ambiguous; keep it covered.
    assert.equal(formatShortcut('Ctrl+-', 'MacIntel'), '⌘-');
  });

  it('returns empty string for empty input (defensive)', () => {
    assert.equal(formatShortcut('', 'Win32'), '');
  });
});
