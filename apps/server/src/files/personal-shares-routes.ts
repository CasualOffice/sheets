import type { FastifyInstance, FastifyReply, FastifyRequest } from 'fastify';
import type {
  FileRecord,
  MemberAcl,
  MemberAclView,
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
 *
 * Member ACL routes (sharing-model §6.2 — MULTI MODE ONLY; single mode
 * is link-only per §8 Q1, so these 404 there). Also the SAFE
 * FOUNDATION: an ACL row grants NO access until the enforcement batch
 * wires `getMemberRole` into the join path.
 *   GET    /files/:id/shares/members
 *   POST   /files/:id/shares/member
 *   PATCH  /files/:id/shares/member/:memberId
 *   DELETE /files/:id/shares/member/:memberId
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

  // ── GET /files/shares/link/:token/meta ──────────────────────────────
  // PUBLIC pre-join discovery (sharing-model §6.1). The token IS the
  // capability, so this is intentionally NOT owner-gated — it's mounted
  // OUTSIDE `ownedFileCtx`. A joiner who holds a `?share=<token>` needs
  // to know, BEFORE opening the collab WS, whether a password is needed
  // (so the client can prompt for `?sp=`) or whether the link is dead.
  //
  // Contract:
  //   - valid, non-expired token → { valid: true, role, hasPassword, roomId }
  //   - unknown / expired token  → { valid: false } and NOTHING else.
  //
  // `getLinkRole` bakes expiry into a null result and surfaces only
  // `hasPassword` (a boolean) — the bcrypt `passwordHash` NEVER leaves
  // the store, so it can't leak here. We deliberately do NOT distinguish
  // unknown from expired in the response (both → { valid: false }) so a
  // probe can't learn that a token ever existed.
  app.get<{ Params: { token: string } }>('/files/shares/link/:token/meta', async (req, reply) => {
    if (store.mode === 'none') {
      // Personal store isn't even instantiated in anonymous-only
      // deploys; mirror the rest of the surface with a 503 so the
      // client can tell "feature off" from "bad token".
      return reply.code(503).send({ error: 'personal-mode-disabled' });
    }
    const link = store.getLinkRole(req.params.token);
    if (!link) {
      return reply.send({ valid: false });
    }
    return reply.send({
      valid: true,
      role: link.role,
      hasPassword: link.hasPassword,
      roomId: link.roomId,
    });
  });

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

  // ── Member ACL routes (sharing-model §6.2 — MULTI MODE ONLY) ─────────
  //
  // Single mode is link-only (§8 Q1: single = one account), so these
  // routes 404 there — `ownedFileCtx` with `{ multiOnly: true }` shadows
  // them. INERT FOUNDATION: an ACL row grants no access until the
  // enforcement batch reads getMemberRole from the join path.

  // ── GET /files/:id/shares/members ───────────────────────────────────
  // List ACL rows (with username/email + role) for a file.
  app.get<{ Params: { id: string } }>('/files/:id/shares/members', async (req, reply) => {
    const ctx = ownedFileCtx(req, reply, store, { multiOnly: true });
    if (!ctx) return;
    const members = store.listMemberAcls(ctx.record.id);
    return reply.send({ members: members.map(toPublicMember) });
  });

  // ── POST /files/:id/shares/member ───────────────────────────────────
  // Grant a member a role. Body: { handle, role } where handle is an
  // email or username (resolved via findMemberByHandle). Refuses to add
  // yourself (the owner is implicitly admin) and the file owner.
  app.post<{
    Params: { id: string };
    Body: { handle?: unknown; role?: unknown };
  }>('/files/:id/shares/member', async (req, reply) => {
    const ctx = ownedFileCtx(req, reply, store, { multiOnly: true });
    if (!ctx) return;
    const body = (req.body ?? {}) as { handle?: unknown; role?: unknown };

    if (typeof body.handle !== 'string' || body.handle.trim().length === 0) {
      return reply.code(400).send({ error: 'invalid-handle' });
    }
    if (!isShareRole(body.role)) {
      return reply.code(400).send({ error: 'invalid-role' });
    }

    const member = store.findMemberByHandle(body.handle);
    if (!member) return reply.code(404).send({ error: 'user-not-found' });

    // The owner already has full access (and an admin is implicitly so).
    // An ACL granting them a lesser role would be confusing + pointless;
    // reject rather than silently shadow it. Also blocks adding yourself.
    if (member.id === ctx.user.id || member.id === ctx.record.ownerId) {
      return reply.code(400).send({ error: 'cannot-add-owner' });
    }

    const acl = store.setMemberAcl({
      workbookId: ctx.record.id,
      memberId: member.id,
      role: body.role,
      createdBy: ctx.user.id,
    });
    return reply.code(201).send(
      toPublicMember({
        ...acl,
        username: member.username,
        email: member.email,
      }),
    );
  });

  // ── PATCH /files/:id/shares/member/:memberId ────────────────────────
  // Change a member's role. Body: { role }.
  app.patch<{
    Params: { id: string; memberId: string };
    Body: { role?: unknown };
  }>('/files/:id/shares/member/:memberId', async (req, reply) => {
    const ctx = ownedFileCtx(req, reply, store, { multiOnly: true });
    if (!ctx) return;
    const memberId = parseMemberId(req.params.memberId, reply);
    if (memberId === INVALID) return;

    const body = (req.body ?? {}) as { role?: unknown };
    if (!isShareRole(body.role)) {
      return reply.code(400).send({ error: 'invalid-role' });
    }
    // Confirm the ACL exists on THIS file before touching it — otherwise
    // an upsert would silently create a row for a member who was never
    // granted (and we have no display info for them here).
    if (store.getMemberRole(ctx.record.id, memberId) === null) {
      return reply.code(404).send({ error: 'not-found' });
    }
    store.setMemberAcl({
      workbookId: ctx.record.id,
      memberId,
      role: body.role,
      createdBy: ctx.user.id,
    });
    const updated = store.listMemberAcls(ctx.record.id).find((m) => m.memberId === memberId);
    if (!updated) return reply.code(404).send({ error: 'not-found' });
    return reply.send(toPublicMember(updated));
  });

  // ── DELETE /files/:id/shares/member/:memberId ───────────────────────
  app.delete<{ Params: { id: string; memberId: string } }>(
    '/files/:id/shares/member/:memberId',
    async (req, reply) => {
      const ctx = ownedFileCtx(req, reply, store, { multiOnly: true });
      if (!ctx) return;
      const memberId = parseMemberId(req.params.memberId, reply);
      if (memberId === INVALID) return;
      if (!store.deleteMemberAcl(ctx.record.id, memberId)) {
        return reply.code(404).send({ error: 'not-found' });
      }
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

/** Public projection of a member ACL — surfaces the display fields the
 *  share dialog needs (username/email/role) without leaking internals. */
function toPublicMember(member: MemberAcl & Pick<MemberAclView, 'username' | 'email'>) {
  return {
    memberId: member.memberId,
    username: member.username,
    email: member.email,
    role: member.role,
    createdAt: member.createdAt,
    createdBy: member.createdBy,
  };
}

/** Parse + validate a `:memberId` path param into a positive integer.
 *  Sends 400 + returns INVALID on a non-numeric / non-positive value. */
function parseMemberId(raw: string, reply: FastifyReply): number | typeof INVALID {
  const id = Number(raw);
  if (!Number.isInteger(id) || id <= 0) {
    reply.code(400).send({ error: 'invalid-member' });
    return INVALID;
  }
  return id;
}

/** Resolve the signed-in user + the owned (or admin-reachable) file,
 *  or send the right error and return null. Mirrors the files-routes
 *  `requireUser` + `ownedFileOr403` pair in one shot. With
 *  `{ multiOnly: true }` the route 404s outside multi mode (member ACLs
 *  are multi-mode only per sharing-model §8 Q1). */
function ownedFileCtx(
  req: FastifyRequest<{ Params: { id: string } }>,
  reply: FastifyReply,
  store: PersonalAuthStore,
  opts: { multiOnly?: boolean } = {},
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
  // Member ACLs are multi-mode only (sharing-model §8 Q1: single = one
  // account, link-only). 404 in single mode — never confirm the file
  // even exists on a route that doesn't apply to the mode.
  if (opts.multiOnly && store.mode !== 'multi') {
    reply.code(404).send({ error: 'not-found' });
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
