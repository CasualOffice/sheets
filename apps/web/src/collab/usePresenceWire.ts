import { useEffect, useMemo, useState } from 'react';
import type { HocuspocusProvider } from '@hocuspocus/provider';
import type { FUniver } from '@univerjs/core/facade';
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

  // Broadcast our selection. We tried subscribing to FUniver's
  // SelectionChanged event but it doesn't fire on programmatic
  // `range.activate()` (only on user-driven moves), which broke parity
  // between scripted and interactive usage. A lightweight 150 ms poll
  // covers both — diff-guarded so we only write when the selection
  // actually changes, which yields one awareness update per cell move
  // in practice.
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
        const next: PeerAwareness['sel'] = {
          u: wb.getId(),
          s: sheetId,
          r: {
            sr: raw.startRow,
            er: raw.endRow,
            sc: raw.startColumn,
            ec: raw.endColumn,
          },
        };
        if (
          prev.sel &&
          prev.sel.u === next.u &&
          prev.sel.s === next.s &&
          prev.sel.r.sr === next.r.sr &&
          prev.sel.r.er === next.r.er &&
          prev.sel.r.sc === next.r.sc &&
          prev.sel.r.ec === next.r.ec
        ) {
          return;
        }
        awareness.setLocalState({ ...prev, sel: next });
      } catch (err) {
        console.warn('[presence] failed to read active selection', err);
      }
    };

    writeIfChanged();
    const id = setInterval(writeIfChanged, 150);
    return () => clearInterval(id);
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
    const onStarted = (p: { row: number; column: number }) => {
      // Reset the per-edit text tracker so an empty-cell start doesn't
      // inherit the previous edit's last value.
      lastText = '';
      writeLive(p.row, p.column, '');
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

  return { peers, myClientId };
}
