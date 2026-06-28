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
 * `formatShortcut('Ctrl+X', navigator.platform)` → `'⌘X'` on Mac,
 * `'Ctrl+X'` on Win/Linux. The canonical form stays `'Ctrl+...'` so
 * callers don't have to detect the platform themselves; this helper
 * is the single conversion site.
 *
 * Why a util:
 * - The shortcuts dialog (Phase 3 #12) needs Mac-correct rendering.
 * - MenuBar currently displays the raw `Ctrl+X` literal even on Mac,
 *   which is wrong. New code should use this util; existing debt
 *   gets paid down on the next touch (memory `feedback-shortcut-strings`).
 *
 * Test-friendly: the platform string is passed in (not read from
 * navigator) so unit tests can simulate both surfaces without
 * monkey-patching globals.
 */

const MAC_SYMBOLS: Record<string, string> = {
  Ctrl: '⌘',
  Cmd: '⌘',
  Alt: '⌥',
  Shift: '⇧',
  // No-op for keys that don't have a Mac symbol — they pass through.
};

const KEY_SYMBOLS: Record<string, string> = {
  PageUp: 'PgUp',
  PageDown: 'PgDn',
  Home: 'Home',
  End: 'End',
  Space: 'Space',
  Escape: 'Esc',
  // `+` doubles as the separator, so callers pass it as `Plus` (or the
  // `++` tail, normalized below) and we render the literal key.
  Plus: '+',
};

/** True when the running platform reports any Mac-flavoured string —
 *  including the Playwright/Chromium quirk that says `Win32` even on
 *  macOS for navigator.platform. Callers that need a hard Mac check
 *  should use this util instead of comparing strings directly. */
export function isMacPlatform(platform: string): boolean {
  // navigator.platform on Mac yields 'MacIntel' on Intel, 'iPhone',
  // 'iPad' etc. Some browsers also expose the deprecated 'Mac68K',
  // 'MacPPC'. Test for any 'Mac', 'iPhone', or 'iPad' substring; the
  // false-positive surface (a desktop env named after a Mac variant)
  // is theoretical.
  return /Mac|iPhone|iPad/.test(platform);
}

/**
 * Convert the canonical `Ctrl+X+...` form to the right rendering for
 * the running platform. Mac collapses `Ctrl+X` into `⌘X` (no plus);
 * Win/Linux keeps the `Ctrl+X` literal with plus separators.
 *
 *   formatShortcut('Ctrl+Shift+V', 'MacIntel') === '⇧⌘V'
 *   formatShortcut('Ctrl+Shift+V', 'Win32')    === 'Ctrl+Shift+V'
 *   formatShortcut('F2', 'MacIntel')           === 'F2'
 *
 * Mac convention is Shift before Cmd before Alt (Apple HIG), so we
 * reorder accordingly. Plain function-keys / F-keys / single keys
 * pass through unchanged.
 */
export function formatShortcut(canonical: string, platform: string): string {
  if (!canonical) return '';
  const isMac = isMacPlatform(platform);
  // A trailing `++` means the key itself is `+` (e.g. Ctrl++ = insert cells).
  // Rewrite to `+Plus` so the split below doesn't swallow it as a separator.
  const normalized = canonical.replace(/\+\+$/, '+Plus');
  const parts = normalized
    .split('+')
    .map((p) => p.trim())
    .filter(Boolean);
  if (parts.length === 0) return canonical;

  // Substitute friendly key names (PgUp / PgDn / Esc).
  const renderKey = (key: string): string => KEY_SYMBOLS[key] ?? key;

  if (!isMac) {
    return parts.map(renderKey).join('+');
  }

  // Mac path: split modifiers from the trailing key. Modifiers stack
  // in Apple-HIG order (Ctrl, Alt/Opt, Shift, Cmd — but render as
  // ⌃⌥⇧⌘) and the final key sits to the right of the symbols.
  const MOD_ORDER: Array<'Ctrl' | 'Alt' | 'Shift' | 'Cmd'> = ['Ctrl', 'Alt', 'Shift', 'Cmd'];
  const presentMods = new Set<string>();
  let key: string | null = null;
  for (const p of parts) {
    if (p === 'Ctrl' || p === 'Cmd' || p === 'Alt' || p === 'Shift') {
      presentMods.add(p);
    } else {
      key = p;
    }
  }
  // Canonical strings use `Ctrl` to mean "primary modifier" — on Mac
  // that's ⌘. If both `Ctrl` and `Cmd` appear we render both, but the
  // typical case is the former-collapses-to-the-latter.
  if (presentMods.has('Ctrl') && !presentMods.has('Cmd')) {
    presentMods.delete('Ctrl');
    presentMods.add('Cmd');
  }
  const modsRendered = MOD_ORDER.filter((m) => presentMods.has(m))
    .map((m) => MAC_SYMBOLS[m] ?? m)
    .join('');
  return modsRendered + (key ? renderKey(key) : '');
}
