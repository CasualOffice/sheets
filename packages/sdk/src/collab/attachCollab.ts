/**
 * attachCollab — opt-in real-time co-editing for `<CasualSheets>`.
 *
 * The editor ships **collab-unaware**: it boots, edits, and persists with no
 * knowledge of rooms or peers. A host that wants co-editing calls
 * `attachCollab(api, { room, server })` once after `onReady` and gets back a
 * detach handle. That's the entire public surface — everything else (presence
 * UI, password prompts, room preflight, reconnect banners) is the host's to
 * build on top of the returned `provider` / `doc`.
 *
 * Transport: Yjs + Hocuspocus, exactly as the reference host (`apps/web`) uses
 * it. The non-negotiable Univer hooks live in `./bridge` — it subscribes to
 * `ICommandService.onMutationExecutedForCollab`, applies remote mutations with
 * `IExecutionOptions.fromCollab` (echo-loop prevention), and guards
 * `params.__splitChunk__`.
 *
 * Persistence note: Yjs/Hocuspocus is the **realtime transport only**. The
 * authoritative document is still saved by the host via WOPI / its own backend
 * (the save/exit event contract) — collab does not turn the SDK into a store.
 */

import * as Y from 'yjs';
import { HocuspocusProvider, HocuspocusProviderWebsocket } from '@hocuspocus/provider';
import type { IWorkbookData } from '@univerjs/core';
import type { CasualSheetsAPI } from '../sheets/api';
import { startBridge, type BridgeHandle } from './bridge';
import { buildWsUrl } from './ws-url';

/** `write` peers broadcast their edits; `view` peers only receive. The
 *  client-side gate is belt-and-braces — real enforcement is the server's
 *  `role` check on the WS upgrade. */
export type CollabRole = 'view' | 'write';

/** Coarse connection state mapped from Hocuspocus's provider status. */
export type CollabConnectionStatus = 'connecting' | 'live' | 'offline';

export interface AttachCollabOptions {
  /** Room / document id. Becomes the Hocuspocus document name. */
  room: string;
  /** Base WebSocket URL of the collab server, e.g. `wss://host/yjs`. */
  server: string;
  /** Room password, if the server gates the room. Sent on the WS URL as
   *  `p=…` so the upgrade handler can validate before the protocol handshake. */
  password?: string;
  /** Auth token for the Hocuspocus handshake. The provider only sends its
   *  auth submessage when this is truthy; servers with an `onAuthenticate`
   *  hook keep the connection queued without it. Defaults to `'anon'` (the
   *  reference server's hook only reads the `role` query param). */
  token?: string;
  /** `view` joins read-only. Defaults to `'write'`. */
  role?: CollabRole;
  /**
   * Called when a peer's compaction snapshot arrives — the host swaps the
   * workbook (typically `api.loadSnapshot(wb)`). MAY return a promise; the
   * bridge pauses op-log replay until it resolves so later mutations don't
   * land on the pre-swap unit.
   */
  onSnapshot?: (wb: IWorkbookData) => void | Promise<void>;
  /** Connection-status transitions — drive a status pill / offline banner. */
  onStatus?: (status: CollabConnectionStatus) => void;
}

export interface CollabHandle {
  /** The underlying Yjs document — introspection / devtools / extra maps. */
  readonly doc: Y.Doc;
  /** The Hocuspocus provider — `awareness` for presence, `on('status')`, etc. */
  readonly provider: HocuspocusProvider;
  /** The mutation bridge — replay-failure diagnostics live here. */
  readonly bridge: BridgeHandle;
  /** Last known connection status. */
  status(): CollabConnectionStatus;
  /** Detach: tear down bridge + provider + doc. Idempotent. */
  detach(): void;
}

/**
 * Attach real-time collab to a live editor. Call once after `onReady`.
 * Returns a {@link CollabHandle}; call `.detach()` to leave the room
 * (and always before the editor unmounts).
 */
export function attachCollab(api: CasualSheetsAPI, opts: AttachCollabOptions): CollabHandle {
  const role: CollabRole = opts.role ?? 'write';

  const doc = new Y.Doc();
  // Match the reference host: drop the reconnect timeout to 10 s so a dropped
  // socket surfaces as `offline` quickly rather than after the 30 s default.
  const ws = new HocuspocusProviderWebsocket({
    url: buildWsUrl(opts.server, opts.room, role, opts.password),
    messageReconnectTimeout: 10_000,
  });
  const provider = new HocuspocusProvider({
    websocketProvider: ws,
    name: opts.room,
    document: doc,
    // Truthy token so the handshake completes even when the server has an
    // onAuthenticate hook (used for role enforcement). See option docs above.
    token: opts.token ?? 'anon',
  });

  const bridge = startBridge(api.univer, doc, {
    role,
    awareness: provider.awareness ?? undefined,
    onSnapshotReceived: opts.onSnapshot,
  });

  let current: CollabConnectionStatus = 'connecting';
  const onStatus = (ev: { status: string }) => {
    const next: CollabConnectionStatus =
      ev.status === 'connected' ? 'live' : ev.status === 'connecting' ? 'connecting' : 'offline';
    if (next === current) return;
    current = next;
    opts.onStatus?.(next);
  };
  provider.on('status', onStatus);

  let detached = false;
  const detach = () => {
    if (detached) return;
    detached = true;
    provider.off('status', onStatus);
    bridge.dispose();
    provider.destroy();
    doc.destroy();
  };

  return {
    doc,
    provider,
    bridge,
    status: () => current,
    detach,
  };
}
