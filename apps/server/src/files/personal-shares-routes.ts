import type { FastifyInstance, FastifyReply, FastifyRequest } from 'fastify';
import type {
  FileRecord,
  PersonalAuthStore,
  PublicUser,
  ShareLink,
  ShareRole,
} from '../auth/personal.js';
import { isShareRole } from '../auth/personal.js';
import { currentUser } from '../auth/personal-routes.js';

/**
 * Personal-mode share-link routes (sharing-model §6.1 — the SAFE
 * FOUNDATION). CRUD over the `share_links` table, gated on file
 * ownership exactly like `personal-files-routes.ts`.
 *
 * Tokens are bound to a specific collab room (`roomId`) at mint time.
 * Enforcement lives in the collab gate (`yjs.ts` onAuthenticate →
 * `resolveJoinRole`), which validates a `?share=<token>` join against
 * the persisted token's room + role + optional password. Rooms are
 * anonymous and NOT keyed by workbookId, so without the room binding a
 * token could be replayed against any room.
 *
 * Gating (matches the files routes):
 *   - owner  = file registry `ownerId === user.id`
 *   - admin  = `user.isAdmin` (cross-user, per §4 RequireAdmin)
 *   - anyone else gets 404 — never leak whether the file exists.
 *
 * Routes (mounted in both single + multi mode; single mode is
 * link-only, which is exactly this surface):
 *   GET    /files/:id/shares
 *   POST   /files/:id/shares/link
 *   PATCH  /files/:id/shares/link/:token
 *   DELETE /files/:id/shares/link/:token
 */

const DAY_MS = 24 * 60 * 60 * 1000;
/** ~273 years — a positive sanity cap so a fat-fingered expiry can't
 *  overflow into a meaningless date. */
const MAX_EXPIRES_DAYS = 100_000;
const MAX_PASSWORD_LEN = 256;
/** Defensive cap on the room id — real ids are 12-char base36, but a
 *  caller could post anything. Bound it so a junk value can't bloat the
 *  row. */
const MAX_ROOM_ID_LEN = 128;

export type SharesRoutesOptions = {
  /** Validates that a posted roomId refers to a live room. When omitted
   *  (e.g. a test that doesn't wire the registry), the roomId is accepted
   *  opaquely — it's still bound to the token and the collab gate's
   *  room-match check is the real enforcement. */
  roomExists?: (roomId: string) => boolean;
};

export function registerPersonalSharesRoutes(
  app: FastifyInstance,
  store: PersonalAuthStore,
  options: SharesRoutesOptions = {},
): void {
  const roomExists = options.roomExists ?? null;
  // ── GET /files/:id/shares ───────────────────────────────────────────
  // List link tokens for a file. passwordHash is never returned — the
  // response carries `hasPassword` instead.
  app.get<{ Params: { id: string } }>('/files/:id/shares', async (req, reply) => {
    const ctx = ownedFileCtx(req, reply, store);
    if (!ctx) return;
    const links = store.listShareLinks(ctx.record.id);
    return reply.send({ links: links.map(toPublicLink) });
  });

  // ── POST /files/:id/shares/link ─────────────────────────────────────
  // Mint a token bound to a specific collab room. Body:
  //   { roomId, role, expiresInDays?, password? }.
  // roomId is REQUIRED — the token is the capability to JOIN that room,
  // and binding it at mint time is what stops a token from being replayed
  // against another room (rooms are anonymous + not keyed by workbookId).
  app.post<{
    Params: { id: string };
    Body: { roomId?: unknown; role?: unknown; expiresInDays?: unknown; password?: unknown };
  }>('/files/:id/shares/link', async (req, reply) => {
    const ctx = ownedFileCtx(req, reply, store);
    if (!ctx) return;
    const body = (req.body ?? {}) as {
      roomId?: unknown;
      role?: unknown;
      expiresInDays?: unknown;
      password?: unknown;
    };

    const roomId = parseRoomId(body.roomId, reply, roomExists);
    if (roomId === INVALID) return;
    if (!isShareRole(body.role)) {
      return reply.code(400).send({ error: 'invalid-role' });
    }
    const expiresAt = parseExpiry(body.expiresInDays, reply);
    if (expiresAt === INVALID) return;
    const password = parsePassword(body.password, reply);
    if (password === INVALID) return;

    const link = store.createShareLink({
      workbookId: ctx.record.id,
      roomId,
      role: body.role,
      createdBy: ctx.user.id,
      expiresAt,
      password,
    });
    // The URL shape is a fragment query the client appends to whatever
    // room/file URL it's already on — we don't hardcode an origin here
    // (the host owns its public URL). See sharing-model §3.5.
    return reply.code(201).send({
      token: link.token,
      roomId: link.roomId,
      role: link.role,
      expiresAt: link.expiresAt,
      url: `?share=${link.token}`,
    });
  });

  // ── PATCH /files/:id/shares/link/:token ─────────────────────────────
  // Flip role and/or expiry. Body: { role?, expiresInDays? }.
  app.patch<{
    Params: { id: string; token: string };
    Body: { role?: unknown; expiresInDays?: unknown };
  }>('/files/:id/shares/link/:token', async (req, reply) => {
    const ctx = ownedFileCtx(req, reply, store);
    if (!ctx) return;
    const link = tokenOnFileOr404(store, ctx.record.id, req.params.token, reply);
    if (!link) return;

    const body = (req.body ?? {}) as { role?: unknown; expiresInDays?: unknown };
    const patch: { role?: ShareRole; expiresAt?: number | null } = {};
    if (body.role !== undefined) {
      if (!isShareRole(body.role)) return reply.code(400).send({ error: 'invalid-role' });
      patch.role = body.role;
    }
    if (body.expiresInDays !== undefined) {
      const expiresAt = parseExpiry(body.expiresInDays, reply);
      if (expiresAt === INVALID) return;
      patch.expiresAt = expiresAt;
    }

    const updated = store.updateShareLink(link.token, patch);
    if (!updated) return reply.code(404).send({ error: 'not-found' });
    return reply.send(toPublicLink(updated));
  });

  // ── DELETE /files/:id/shares/link/:token ────────────────────────────
  app.delete<{ Params: { id: string; token: string } }>(
    '/files/:id/shares/link/:token',
    async (req, reply) => {
      const ctx = ownedFileCtx(req, reply, store);
      if (!ctx) return;
      const link = tokenOnFileOr404(store, ctx.record.id, req.params.token, reply);
      if (!link) return;
      store.deleteShareLink(link.token);
      return reply.code(204).send();
    },
  );
}

// ── Helpers ──────────────────────────────────────────────────────────

/** Sentinel returned by the parse helpers when they've already sent a
 *  400 — the caller bails without sending a second reply. */
const INVALID = Symbol('invalid');

/** Public projection of a link — drops passwordHash, exposes a
 *  boolean instead. */
function toPublicLink(link: ShareLink) {
  return {
    token: link.token,
    roomId: link.roomId,
    role: link.role,
    expiresAt: link.expiresAt,
    hasPassword: link.passwordHash !== null,
    createdAt: link.createdAt,
    createdBy: link.createdBy,
  };
}

/** Resolve the signed-in user + the owned (or admin-reachable) file,
 *  or send the right error and return null. Mirrors the files-routes
 *  `requireUser` + `ownedFileOr403` pair in one shot. */
function ownedFileCtx(
  req: FastifyRequest<{ Params: { id: string } }>,
  reply: FastifyReply,
  store: PersonalAuthStore,
): { user: PublicUser; record: FileRecord } | null {
  if (store.mode === 'none') {
    reply.code(503).send({ error: 'personal-mode-disabled' });
    return null;
  }
  const user = currentUser(req, store);
  if (!user) {
    reply.code(401).send({ error: 'unauthenticated' });
    return null;
  }
  const record = store.getFile(req.params.id);
  // 404 (not 403) for both unknown + non-owner — don't leak existence.
  // Admins additionally pass for any file (§4 RequireAdmin).
  if (!record || (record.ownerId !== user.id && !user.isAdmin)) {
    reply.code(404).send({ error: 'not-found' });
    return null;
  }
  return { user, record };
}

/** Load a token and confirm it belongs to this file — 404 otherwise so
 *  a token from another workbook can't be edited via this file's path. */
function tokenOnFileOr404(
  store: PersonalAuthStore,
  workbookId: string,
  token: string,
  reply: FastifyReply,
): ShareLink | null {
  const link = store.getShareLink(token);
  if (!link || link.workbookId !== workbookId) {
    reply.code(404).send({ error: 'not-found' });
    return null;
  }
  return link;
}

/** Parse the optional `expiresInDays` into an absolute ms epoch (or
 *  null when omitted / explicitly null). Sends 400 + returns INVALID
 *  on a non-positive / out-of-range value. */
function parseExpiry(raw: unknown, reply: FastifyReply): number | null | typeof INVALID {
  if (raw === undefined || raw === null) return null;
  if (typeof raw !== 'number' || !Number.isFinite(raw) || raw <= 0 || raw > MAX_EXPIRES_DAYS) {
    reply.code(400).send({ error: 'invalid-expiry' });
    return INVALID;
  }
  return Date.now() + raw * DAY_MS;
}

/** Validate the REQUIRED roomId. Must be a non-empty, sanely-bounded
 *  string. When a `roomExists` checker is supplied, the room must be
 *  live — a 404 (not 400) so we don't confirm whether a guessed room id
 *  exists to a caller who shouldn't know. Sends the error + returns
 *  INVALID on failure. */
function parseRoomId(
  raw: unknown,
  reply: FastifyReply,
  roomExists: ((roomId: string) => boolean) | null,
): string | typeof INVALID {
  if (typeof raw !== 'string') {
    reply.code(400).send({ error: 'invalid-room' });
    return INVALID;
  }
  const trimmed = raw.trim();
  if (trimmed.length === 0 || trimmed.length > MAX_ROOM_ID_LEN) {
    reply.code(400).send({ error: 'invalid-room' });
    return INVALID;
  }
  if (roomExists && !roomExists(trimmed)) {
    reply.code(404).send({ error: 'room-not-found' });
    return INVALID;
  }
  return trimmed;
}

/** Validate the optional join password. Empty / omitted → no password
 *  (null). Sends 400 + returns INVALID on a non-string / over-long. */
function parsePassword(raw: unknown, reply: FastifyReply): string | null | typeof INVALID {
  if (raw === undefined || raw === null) return null;
  if (typeof raw !== 'string' || raw.length > MAX_PASSWORD_LEN) {
    reply.code(400).send({ error: 'invalid-password' });
    return INVALID;
  }
  const trimmed = raw.trim();
  return trimmed.length === 0 ? null : trimmed;
}
