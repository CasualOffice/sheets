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
 * Local-user identity for the presence layer. Persisted in localStorage
 * so a name set in one tab follows the user across reloads and into new
 * rooms. Color is deterministically derived from the name so a single
 * user picks the *same* color in every room — useful when two tabs of
 * the same person open the same room.
 */

const NAME_KEY = 'casual.collab.displayName';
const NAME_PROMPTED_KEY = 'casual.collab.namePrompted';

export type Identity = {
  name: string;
  color: string;
};

/**
 * Palette tuned for AA contrast on white backgrounds with white text.
 * Twelve hues so even a packed room rarely collides.
 */
const COLORS = [
  '#1a73e8', '#d93025', '#188038', '#e8710a', '#9334e6', '#1e8e3e',
  '#c5221f', '#1967d2', '#a142f4', '#f29900', '#0b8043', '#b06000',
];

export function colorForName(name: string): string {
  const trimmed = name.trim() || 'anon';
  let h = 5381;
  for (let i = 0; i < trimmed.length; i += 1) {
    // djb2 — fast, well-distributed for short strings.
    h = ((h << 5) + h + trimmed.charCodeAt(i)) | 0;
  }
  return COLORS[Math.abs(h) % COLORS.length];
}

export function getDisplayName(): string | null {
  try {
    const v = localStorage.getItem(NAME_KEY);
    return v && v.trim().length > 0 ? v : null;
  } catch {
    return null;
  }
}

export function setDisplayName(name: string): void {
  const clean = name.trim().slice(0, 32);
  try {
    localStorage.setItem(NAME_KEY, clean);
    localStorage.setItem(NAME_PROMPTED_KEY, '1');
  } catch {
    /* private mode — fine, we'll just re-prompt */
  }
}

export function markNamePrompted(): void {
  try {
    localStorage.setItem(NAME_PROMPTED_KEY, '1');
  } catch {
    /* no-op */
  }
}

export function wasNamePrompted(): boolean {
  try {
    return localStorage.getItem(NAME_PROMPTED_KEY) === '1';
  } catch {
    return false;
  }
}

/**
 * Whimsical default if the user dismisses the prompt — better than
 * "Anonymous" in a room with multiple unnamed peers, and the seed
 * keeps it stable per browser session.
 */
const ADJ = ['Curious', 'Sunny', 'Quiet', 'Brisk', 'Cosmic', 'Quirky', 'Gentle', 'Lively'];
const ANIM = ['Otter', 'Sparrow', 'Fox', 'Panda', 'Heron', 'Lynx', 'Badger', 'Magpie'];

export function suggestAnonName(): string {
  const a = ADJ[Math.floor(Math.random() * ADJ.length)];
  const b = ANIM[Math.floor(Math.random() * ANIM.length)];
  return `${a} ${b}`;
}

export function initials(name: string): string {
  const parts = name.trim().split(/\s+/).filter(Boolean);
  if (parts.length === 0) return '?';
  if (parts.length === 1) return parts[0].slice(0, 2).toUpperCase();
  return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
}

/**
 * Awareness payload broadcast per peer. Kept tiny — awareness updates
 * fire on every selection move and ride the same WS as Yjs sync.
 *
 *   u: unitId of the workbook the selection belongs to (peers ignore
 *      selections from a different unit — usually our local active one
 *      remapped, but during snapshot swaps they can diverge briefly).
 *   s: subUnitId (sheet id) — peer's selection only renders when the
 *      local active sheet matches.
 *   r: bounding rect {sr, er, sc, ec} of the primary selection range.
 *   rs: optional array of ALL ranges in a multi-range (Ctrl-click)
 *      selection. Includes the primary at index 0 by convention so
 *      the presence layer can iterate `rs` when present and fall back
 *      to `r` for legacy peers. Older clients ignore `rs` and keep
 *      seeing just the primary — backwards compatible.
 */
export type PeerAwareness = {
  name: string;
  color: string;
  sel?: {
    u: string;
    s: string;
    r: { sr: number; er: number; sc: number; ec: number };
    rs?: Array<{ sr: number; er: number; sc: number; ec: number }>;
  };
  /** In-progress cell edit. Cleared on SheetEditEnded. Used to render a
   *  "ghost" overlay so peers see the value appearing character-by-
   *  character instead of jumping in on commit. */
  liveEdit?: {
    u: string;
    s: string;
    row: number;
    col: number;
    text: string;
  };
  /** Hex-encoded Y.Doc state vector. Broadcast every few seconds so
   *  peers can compare and detect a stuck sync. Same SV means peers
   *  have seen the same op-log entries (≠ same applied state, but
   *  catches the most common divergence: Yjs sync stalled). */
  sv?: string;
  /** Unix ms when `sv` was last computed. Used by the divergence
   *  detector to age out stale readings — a peer that hasn't updated
   *  in 30 s is treated as offline rather than diverged. */
  svAt?: number;
};

export type Peer = {
  /** Yjs awareness clientId — stable per browser tab connection. */
  clientId: number;
  name: string;
  color: string;
  selection: PeerAwareness['sel'];
  liveEdit: PeerAwareness['liveEdit'];
  /** Unix ms of the last awareness update from this peer (from
   *  `awareness.meta[clientId].lastUpdated`). Used by the avatar tooltip
   *  to render "Active now" / "Last seen Xs ago". */
  lastSeen: number;
  /** Latest state-vector hash + capture time. See PeerAwareness.sv. */
  sv?: string;
  svAt?: number;
};
