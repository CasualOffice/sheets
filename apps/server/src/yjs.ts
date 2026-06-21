import { Hocuspocus } from '@hocuspocus/server';
import { WebSocketServer, type WebSocket } from 'ws';
import type { IncomingMessage, Server } from 'node:http';
import * as Y from 'yjs';
import type { RoomRegistry } from './rooms.js';
import type { DocStorage } from './storage.js';
import type { ShareLinkRole } from './auth/personal.js';
import { resolveJoinRole } from './auth/join-role.js';

/** Resolve a share token to its persisted role/room/password, respecting
 *  expiry, or null. In production this is `PersonalAuthStore.getLinkRole`,
 *  bound in index.ts; null in anonymous-only deploys (no personal store),
 *  in which case any `?share=` token is treated as invalid and rejected. */
export type ResolveLinkRole = (token: string) => ShareLinkRole | null;

export type AttachHocuspocusOptions = {
  /** Wired only in personal mode. When absent, share tokens can't be
   *  validated, so a `?share=` join is rejected rather than trusted. */
  resolveLinkRole?: ResolveLinkRole | null;
};

/**
 * Wire Hocuspocus to a Node http server. Hocuspocus owns the Yjs sync
 * protocol; we hook lifecycle callbacks into our RoomRegistry (accounting)
 * and DocStorage (persistence). The room id is the document name — clients
 * connect with `new HocuspocusProvider({ url, name: roomId })`.
 */
export function attachHocuspocus(
  httpServer: Server,
  rooms: RoomRegistry,
  storage: DocStorage,
  pathPrefix = '/yjs',
  options: AttachHocuspocusOptions = {},
): { hocuspocus: Hocuspocus; close: () => Promise<void> } {
  const resolveLinkRole = options.resolveLinkRole ?? null;
  // Debounce per-room saves so a rapid burst of edits doesn't hammer Redis.
  // 500 ms feels right for "still feels live, doesn't write on every keystroke".
  const SAVE_DEBOUNCE_MS = 500;
  const pendingSaves = new Map<string, ReturnType<typeof setTimeout>>();
  const queueSave = (name: string, doc: Y.Doc) => {
    const existing = pendingSaves.get(name);
    if (existing) clearTimeout(existing);
    pendingSaves.set(
      name,
      setTimeout(async () => {
        pendingSaves.delete(name);
        try {
          const update = Y.encodeStateAsUpdate(doc);
          await storage.save(name, update);
        } catch (err) {
          console.warn('[hocuspocus] save failed for', name, err);
        }
      }, SAVE_DEBOUNCE_MS),
    );
  };

  const hocuspocus = new Hocuspocus({
    name: 'casual-sheets',
    async onLoadDocument({ documentName, document }) {
      // 1) If the room was pre-seeded via /api/rooms (xlsx upload path),
      //    apply that first.
      const room = rooms.get(documentName);
      if (room?.seed) Y.applyUpdate(document, room.seed);
      // 2) Restore the latest persisted state from the storage backend.
      const persisted = await storage.load(documentName);
      if (persisted) Y.applyUpdate(document, persisted);
      return document;
    },
    /**
     * Server-side view-only enforcement. Until this hook landed, only
     * the client's `applyViewOnlyMode` + Univer's permission gate
     * stopped a view-role joiner from mutating the room — easily
     * bypassed by a crafted client.
     *
     * Hocuspocus's `ConnectionConfiguration.readOnly` flag, when set
     * in `onAuthenticate`, makes the underlying message receiver
     * reject any incoming sync update for that socket. Setting it
     * here is the only authoritative gate we control.
     *
     * Auth otherwise stays open — the room-password gate runs in the
     * upgrade handler (HTTP-level), and Hocuspocus's own auth flow is
     * effectively a no-op for anonymous rooms.
     *
     * Share-token enforcement (sharing-model §6.1): when the join URL
     * carries `?share=<token>`, the server becomes AUTHORITATIVE. We
     * IGNORE the client `?role=` entirely and derive the privilege from
     * the persisted, room-bound token via the pure `resolveJoinRole`.
     * Any token failure (invalid / expired / wrong room / bad share
     * password) THROWS — Hocuspocus turns a throw in onAuthenticate into
     * a closed connection (its `Unauthorized` close), matching the
     * unauthorized pattern the upgrade handler uses for the room
     * password. Without a token we fall through to the EXACT legacy
     * anonymous behaviour below.
     */
    async onAuthenticate({ requestParameters, connection, documentName }) {
      const token = requestParameters.get('share');

      const decision = resolveJoinRole({
        token: token ?? null,
        documentName,
        sharePassword: requestParameters.get('sp') ?? null,
        // No personal store wired (anonymous-only deploy) → no token can
        // be validated, so treat every token as unknown (rejects below).
        lookup: (t) => (resolveLinkRole ? resolveLinkRole(t) : null),
      });

      if ('reject' in decision) {
        // Refuse the connection. Throwing is the documented way to fail
        // Hocuspocus auth; it closes the socket. We deliberately do NOT
        // leak WHICH check failed to the client (all reasons collapse to
        // one error) — only the server log distinguishes them.
        throw new Error(`share-token rejected: ${decision.reject}`);
      }

      if (decision.via === 'share-token') {
        connection.readOnly = decision.readOnly;
        return { role: decision.role, via: 'share-token' as const };
      }

      // ── No token → legacy anonymous path, unchanged. ────────────────
      const role = requestParameters.get('role');
      if (role === 'view') {
        connection.readOnly = true;
      }
      // Returned object becomes `context` on later hooks. Surface the
      // role so logs / metrics can attribute writes correctly.
      return { role: role === 'view' ? 'view' : 'write' };
    },
    async onChange({ documentName, document }) {
      queueSave(documentName, document as Y.Doc);
    },
    async onConnect({ documentName }) {
      rooms.onConnect(documentName);
    },
    async onDisconnect({ documentName }) {
      rooms.onDisconnect(documentName);
    },
  });

  // Fastify's WebSocket plugin would also work but we keep dependencies
  // minimal — raw ws + manual upgrade routing.
  const wss = new WebSocketServer({ noServer: true });

  const onUpgrade = (req: IncomingMessage, socket: import('node:net').Socket, head: Buffer) => {
    const rawUrl = req.url ?? '/';
    if (!rawUrl.startsWith(pathPrefix)) return;
    // Parse query string so the client can authenticate the WS upgrade
    // with `?room=<id>&p=<password>`. Without the room id we can't
    // validate, so we reject — Hocuspocus would otherwise let the
    // connection through and discover the room name from the protocol
    // handshake AFTER we've already accepted.
    const parsed = new URL(rawUrl, 'http://internal');
    const roomId = parsed.searchParams.get('room');
    const password = parsed.searchParams.get('p');
    const passwordBad =
      roomId !== null && rooms.get(roomId) !== undefined && !rooms.passwordOk(roomId, password);
    // Always complete the WS upgrade — even for bad-password rejections.
    // Pre-upgrade HTTP 401 responses surface in the browser only as
    // close-code 1006 ("abnormal"), indistinguishable from network
    // drops, so the client can't tell auth from outage and silently
    // reconnects forever. Completing the upgrade then closing with the
    // app-defined 4401 code gives the client a frame it can route to
    // the password prompt.
    wss.handleUpgrade(req, socket, head, (ws) => {
      if (passwordBad) {
        try {
          // 4401 = app-defined "unauthorized" (RFC 6455 reserves 4000–4999
          // for private use). The client's onClose handler matches on this.
          ws.close(4401, 'unauthorized');
        } catch {
          // Best-effort — if close throws we just drop the socket.
          try {
            ws.terminate();
          } catch {
            /* swallow */
          }
        }
        return;
      }
      handleConnection(ws, req);
    });
  };

  httpServer.on('upgrade', onUpgrade);

  const handleConnection = (ws: WebSocket, req: IncomingMessage) => {
    // Hocuspocus expects (websocket, request, context). Context is optional
    // and we don't need auth tokens for anonymous rooms.
    hocuspocus.handleConnection(
      ws as unknown as Parameters<typeof hocuspocus.handleConnection>[0],
      req,
    );
  };

  return {
    hocuspocus,
    close: async () => {
      // Flush any pending debounced saves before we tear down.
      for (const [, timer] of pendingSaves) clearTimeout(timer);
      pendingSaves.clear();
      httpServer.off('upgrade', onUpgrade);
      wss.close();
      await hocuspocus.destroy();
    },
  };
}
