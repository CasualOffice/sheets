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

import { useEffect, useMemo, useState } from 'react';
import type { HocuspocusProvider } from '@hocuspocus/provider';
import type { FUniver } from '@univerjs/core/facade';
import * as Y from 'yjs';
import { colorForName, type Identity, type Peer, type PeerAwareness } from './presence';

/**
 * Wires Yjs awareness to Univer selection events:
 *
 *  - Writes our local awareness state {name, color, sel} whenever the
 *    user moves their selection.
 *  - Subscribes to awareness changes from peers and exposes them as a
 *    sorted React state.
 *
 * Returns `[peers, myClientId]` for downstream consumers (the overlay
 * uses the peer list; the avatar stack uses both to deduplicate self).
 *
 * Selection broadcast is debounced via rAF so dragging across cells
 * doesn't flood the WS with awareness updates — Yjs already coalesces
 * but coalescing inside one animation frame is the cheap baseline.
 */
export function usePresenceWire(
  api: FUniver | null,
  provider: HocuspocusProvider | null,
  identity: Identity | null,
): { peers: Peer[]; myClientId: number | null } {
  const [peers, setPeers] = useState<Peer[]>([]);

  const awareness = useMemo(() => provider?.awareness ?? null, [provider]);
  const myClientId = awareness?.clientID ?? null;

  // Push identity into our awareness state whenever it changes.
  useEffect(() => {
    if (!awareness || !identity) return;
    const prev = (awareness.getLocalState() ?? {}) as PeerAwareness;
    awareness.setLocalState({ ...prev, name: identity.name, color: identity.color });
  }, [awareness, identity]);

  // Subscribe to peer-state changes.
  useEffect(() => {
    if (!awareness) {
      setPeers([]);
      return;
    }
    const recompute = () => {
      const out: Peer[] = [];
      // Read per-client metadata for `lastUpdated` (used by AvatarStack's
      // "Active now / Last seen" tooltip). y-protocols/awareness maintains
      // this map alongside the state map and updates lastUpdated on every
      // setLocalState.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const meta = (awareness as unknown as { meta: Map<number, { lastUpdated: number }> }).meta;
      awareness.getStates().forEach((raw, clientId) => {
        if (clientId === awareness.clientID) return; // skip self
        const s = raw as PeerAwareness;
        if (!s || typeof s.name !== 'string') return;
        out.push({
          clientId,
          name: s.name,
          color: typeof s.color === 'string' ? s.color : colorForName(s.name),
          selection: s.sel,
          liveEdit: s.liveEdit,
          lastSeen: meta?.get(clientId)?.lastUpdated ?? Date.now(),
          sv: typeof s.sv === 'string' ? s.sv : undefined,
          svAt: typeof s.svAt === 'number' ? s.svAt : undefined,
        });
      });
      out.sort((a, b) => a.clientId - b.clientId);
      setPeers(out);
    };
    awareness.on('change', recompute);
    // Also tick once a second so "last seen Ns ago" labels refresh even
    // when peers are quiet (no awareness change events firing).
    const interval = setInterval(recompute, 1000);
    recompute();
    return () => {
      awareness.off('change', recompute);
      clearInterval(interval);
    };
  }, [awareness]);

  // Broadcast our selection. Combined strategy:
  //
  //   1. Subscribe to Univer's SelectionChanged event so user-driven
  //      moves (mouse, arrow keys, Tab) fire awareness updates within
  //      one tick — feels instant on the peer side.
  //   2. Keep a lightweight poll as a fallback for programmatic
  //      `range.activate()` calls, which (in Univer 0.22.x at least)
  //      do NOT fire SelectionChanged. The poll's interval is bumped
  //      to 500 ms now that the event covers the hot path — 150 ms
  //      was the wrong knob to compensate for the missing event.
  //
  // Both paths call the same diff-guarded `writeIfChanged` so no
  // matter which fires first, we only ship awareness if the selection
  // actually moved.
  useEffect(() => {
    if (!api || !awareness) return;

    const writeIfChanged = (): void => {
      try {
        const wb = api.getActiveWorkbook();
        if (!wb) return;
        const ws = wb.getActiveSheet();
        if (!ws) return;
        const primary = ws.getActiveRange();
        const prev = (awareness.getLocalState() ?? {}) as PeerAwareness;
        if (!primary) {
          if (prev.sel) awareness.setLocalState({ ...prev, sel: undefined });
          return;
        }
        const raw = primary.getRange();
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const sheetId = (ws as any).getSheetId?.() ?? (ws as any).getId?.() ?? '';
        // Read every range in the current selection so multi-range
        // (Ctrl-click / Shift+F8) selections propagate fully. The
        // facade's `getSelection().getActiveRangeList()` returns
        // FRanges including the primary at index 0; fall back to
        // [primary] if the facade method is missing on older builds.
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const wsAny = ws as any;
        let allRects: Array<{ sr: number; er: number; sc: number; ec: number }>;
        try {
          const sel = wsAny.getSelection?.();
          const list = sel?.getActiveRangeList?.() as
            | Array<{ getRange: () => { startRow: number; endRow: number; startColumn: number; endColumn: number } }>
            | undefined;
          if (list && list.length > 0) {
            allRects = list.map((r) => {
              const g = r.getRange();
              return { sr: g.startRow, er: g.endRow, sc: g.startColumn, ec: g.endColumn };
            });
          } else {
            allRects = [{ sr: raw.startRow, er: raw.endRow, sc: raw.startColumn, ec: raw.endColumn }];
          }
        } catch {
          allRects = [{ sr: raw.startRow, er: raw.endRow, sc: raw.startColumn, ec: raw.endColumn }];
        }
        const next: PeerAwareness['sel'] = {
          u: wb.getId(),
          s: sheetId,
          // Primary kept for legacy peers that don't read `rs`.
          r: allRects[0],
          // Only ship `rs` when there's more than one range — keeps
          // single-range awareness payloads the same size as before.
          rs: allRects.length > 1 ? allRects : undefined,
        };
        if (
          prev.sel &&
          prev.sel.u === next.u &&
          prev.sel.s === next.s &&
          rectsEqual(prev.sel.r, next.r) &&
          rectListsEqual(prev.sel.rs, next.rs)
        ) {
          return;
        }
        awareness.setLocalState({ ...prev, sel: next });
      } catch (err) {
        console.warn('[presence] failed to read active selection', err);
      }
    };
    const rectsEqual = (
      a: { sr: number; er: number; sc: number; ec: number },
      b: { sr: number; er: number; sc: number; ec: number },
    ) => a.sr === b.sr && a.er === b.er && a.sc === b.sc && a.ec === b.ec;
    const rectListsEqual = (
      a?: Array<{ sr: number; er: number; sc: number; ec: number }>,
      b?: Array<{ sr: number; er: number; sc: number; ec: number }>,
    ) => {
      if (!a && !b) return true;
      if (!a || !b) return false;
      if (a.length !== b.length) return false;
      for (let i = 0; i < a.length; i += 1) {
        if (!rectsEqual(a[i], b[i])) return false;
      }
      return true;
    };

    writeIfChanged();
    // Event-driven path: instant updates on user moves.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const anyApi = api as any;
    const SelectionEv = anyApi.Event?.SelectionChanged;
    const eventSub: { dispose?: () => void } | undefined =
      SelectionEv && typeof anyApi.addEvent === 'function'
        ? anyApi.addEvent(SelectionEv, () => writeIfChanged())
        : undefined;
    // Slow-poll fallback for programmatic moves.
    const id = setInterval(writeIfChanged, 500);
    return () => {
      clearInterval(id);
      eventSub?.dispose?.();
    };
  }, [api, awareness]);

  // Live-typing ghost: broadcast in-progress edits via awareness so peers
  // see characters appear in real time, rather than only on commit. We
  // throttle to ~30 ms (one awareness frame per typed key max) so a fast
  // typist doesn't flood the WS.
  useEffect(() => {
    if (!api || !awareness) return;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const anyApi = api as any;
    const ChangingEv = anyApi.Event?.SheetEditChanging;
    const EndedEv = anyApi.Event?.SheetEditEnded;
    const StartedEv = anyApi.Event?.SheetEditStarted;
    if (!ChangingEv || !EndedEv || typeof anyApi.addEvent !== 'function') {
      console.warn('[presence] FUniver missing edit events — version mismatch?');
      return;
    }

    let lastText = '';
    let lastWriteAt = 0;
    let pendingTimer: ReturnType<typeof setTimeout> | null = null;

    const writeLive = (row: number, col: number, text: string): void => {
      const wb = api.getActiveWorkbook();
      if (!wb) return;
      const ws = wb.getActiveSheet();
      if (!ws) return;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const sheetId = (ws as any).getSheetId?.() ?? (ws as any).getId?.() ?? '';
      const prev = (awareness.getLocalState() ?? {}) as PeerAwareness;
      const liveEdit: PeerAwareness['liveEdit'] = {
        u: wb.getId(),
        s: sheetId,
        row,
        col,
        text,
      };
      awareness.setLocalState({ ...prev, liveEdit });
    };

    const clearLive = (): void => {
      const prev = (awareness.getLocalState() ?? {}) as PeerAwareness;
      if (prev.liveEdit) awareness.setLocalState({ ...prev, liveEdit: undefined });
      lastText = '';
      if (pendingTimer) {
        clearTimeout(pendingTimer);
        pendingTimer = null;
      }
    };

    const onChanging = (p: {
      row: number;
      column: number;
      value: { toPlainText?: () => string };
    }) => {
      try {
        const text = typeof p.value?.toPlainText === 'function' ? p.value.toPlainText() : '';
        if (text === lastText) return;
        lastText = text;
        const now = Date.now();
        const elapsed = now - lastWriteAt;
        if (elapsed >= 30) {
          lastWriteAt = now;
          writeLive(p.row, p.column, text);
          if (pendingTimer) {
            clearTimeout(pendingTimer);
            pendingTimer = null;
          }
        } else if (!pendingTimer) {
          // Trailing-edge flush: ensure the final keystroke of a burst
          // also makes it onto the wire even if it landed inside the
          // 30 ms window.
          pendingTimer = setTimeout(() => {
            pendingTimer = null;
            lastWriteAt = Date.now();
            writeLive(p.row, p.column, lastText);
          }, 30 - elapsed);
        }
      } catch (err) {
        console.warn('[presence] live-edit write failed', err);
      }
    };

    const onEnded = () => clearLive();
    const onStarted = (_p: { row: number; column: number }) => {
      // Only reset the per-edit text tracker. Don't write an empty-text
      // liveEdit yet — that caused the ghost to "jump" to the next
      // cell on Tab/Enter for one frame between EditEnded clearing
      // the old position and the user actually starting to type. The
      // first `onChanging` will write the ghost once real text exists.
      lastText = '';
    };

    const subs = [
      anyApi.addEvent(StartedEv, onStarted),
      anyApi.addEvent(ChangingEv, onChanging),
      anyApi.addEvent(EndedEv, onEnded),
    ] as Array<{ dispose?: () => void } | undefined>;

    return () => {
      for (const s of subs) s?.dispose?.();
      if (pendingTimer) clearTimeout(pendingTimer);
      clearLive();
    };
  }, [api, awareness]);

  // Divergence-detection heartbeat: every 5 s, encode our Y.Doc state
  // vector and put it on awareness. Peers compare to their own SV and
  // surface "out of sync" when they disagree for >15 s (see
  // CollabIndicator + collab-context syncHealth). The SV is small (one
  // varint per active clientId) and ships as hex; updating it doesn't
  // tick selection-change UX because we set it on the same awareness
  // state we already maintain.
  useEffect(() => {
    if (!awareness || !provider?.document) return;
    const doc = provider.document;
    const writeSv = () => {
      try {
        const sv = Y.encodeStateVector(doc);
        const hex = bytesToHex(sv);
        const prev = (awareness.getLocalState() ?? {}) as PeerAwareness;
        if (prev.sv === hex) return; // nothing changed since last broadcast
        awareness.setLocalState({ ...prev, sv: hex, svAt: Date.now() });
      } catch (err) {
        console.warn('[presence] failed to encode state vector', err);
      }
    };
    writeSv();
    const id = setInterval(writeSv, 5000);
    return () => clearInterval(id);
  }, [awareness, provider]);

  return { peers, myClientId };
}

function bytesToHex(bytes: Uint8Array): string {
  // Hex is fine — base64 would be ~25% smaller but hex is human-readable
  // in devtools and the payload is tiny anyway (< 32 bytes for typical
  // rooms).
  let out = '';
  for (let i = 0; i < bytes.length; i += 1) {
    const b = bytes[i];
    if (b < 16) out += '0';
    out += b.toString(16);
  }
  return out;
}
