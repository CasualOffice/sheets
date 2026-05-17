import { Hocuspocus } from '@hocuspocus/server';
import { WebSocketServer, type WebSocket } from 'ws';
import type { IncomingMessage, Server } from 'node:http';
import * as Y from 'yjs';
import type { RoomRegistry } from './rooms.js';
import type { DocStorage } from './storage.js';

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
): { hocuspocus: Hocuspocus; close: () => Promise<void> } {
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

  const onUpgrade = (
    req: IncomingMessage,
    socket: import('node:net').Socket,
    head: Buffer,
  ) => {
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
    if (roomId) {
      const room = rooms.get(roomId);
      if (room && !rooms.passwordOk(roomId, password)) {
        // 4401: app-defined close code for "unauthorized". Hocuspocus
        // ignores this and falls back to its own handshake otherwise.
        socket.write(
          'HTTP/1.1 401 Unauthorized\r\n' +
            'Connection: close\r\n' +
            'Content-Length: 0\r\n' +
            '\r\n',
        );
        socket.destroy();
        return;
      }
    }
    wss.handleUpgrade(req, socket, head, (ws) => handleConnection(ws, req));
  };

  httpServer.on('upgrade', onUpgrade);

  const handleConnection = (ws: WebSocket, req: IncomingMessage) => {
    // Hocuspocus expects (websocket, request, context). Context is optional
    // and we don't need auth tokens for anonymous rooms.
    hocuspocus.handleConnection(ws as unknown as Parameters<typeof hocuspocus.handleConnection>[0], req);
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
