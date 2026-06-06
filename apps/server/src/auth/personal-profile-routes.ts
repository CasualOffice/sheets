import type { FastifyInstance, FastifyReply, FastifyRequest } from 'fastify';
import { Readable } from 'node:stream';
import type { PersonalAuthStore, PublicUser, UserProfile } from './personal.js';
import { currentUser } from './personal-routes.js';

/**
 * Profile + preferences routes (Phase C Batch 4.5 of #49).
 *
 *   GET    /auth/profile                       — full profile
 *   PATCH  /auth/profile                       — update fields
 *   POST   /auth/profile/avatar                — upload avatar
 *   DELETE /auth/profile/avatar                — clear avatar
 *   GET    /auth/profile/avatar?id=<userId>    — serve bytes
 *
 * The GET /avatar route accepts a query `id` so the AccountMenu can
 * render avatars for other users without requiring those rows to be
 * exposed in /auth/me. Avatars are public-ish — anyone with a
 * session can fetch any user's avatar, since avatars in a personal
 * docker are essentially display names with pictures. Team mode
 * will tighten this to "members of the same workspace."
 */

const MAX_AVATAR_BYTES = 256 * 1024; // 256 KB cap — UI resizes on upload
const ALLOWED_AVATAR_MIMES = new Set(['image/png', 'image/jpeg', 'image/webp', 'image/gif']);

export function registerPersonalProfileRoutes(
  app: FastifyInstance,
  store: PersonalAuthStore,
): void {
  // ── GET /auth/profile ───────────────────────────────────────────────
  app.get('/auth/profile', async (req, reply) => {
    const user = requireUser(req, reply, store);
    if (!user) return;
    const profile = store.getProfile(user.id);
    if (!profile) return reply.code(404).send({ error: 'not-found' });
    return reply.send({ user, profile });
  });

  // ── PATCH /auth/profile ─────────────────────────────────────────────
  app.patch('/auth/profile', async (req, reply) => {
    const user = requireUser(req, reply, store);
    if (!user) return;
    const body = (req.body ?? {}) as Partial<{
      displayName: string | null;
      email: string | null;
      timezone: string;
      preferences: Record<string, unknown>;
    }>;

    const patch: Parameters<PersonalAuthStore['updateProfile']>[1] = {};
    if ('displayName' in body)
      patch.displayName = typeof body.displayName === 'string' ? body.displayName : null;
    if ('email' in body) patch.email = typeof body.email === 'string' ? body.email : null;
    if ('timezone' in body && typeof body.timezone === 'string') {
      patch.timezone = body.timezone;
    }
    if ('preferences' in body && body.preferences && typeof body.preferences === 'object') {
      patch.preferences = body.preferences as Record<string, unknown>;
    }

    const next = store.updateProfile(user.id, patch);
    if (!next) return reply.code(409).send({ error: 'conflict-or-invalid' });
    return reply.send({ profile: next });
  });

  // ── POST /auth/profile/avatar ───────────────────────────────────────
  app.post('/auth/profile/avatar', async (req, reply) => {
    const user = requireUser(req, reply, store);
    if (!user) return;
    const parts = req.parts();
    let bytes: Buffer | null = null;
    let mime: string | null = null;
    for await (const part of parts) {
      if (part.type === 'file' && part.fieldname === 'avatar') {
        mime = part.mimetype;
        bytes = await streamToBuffer(part.file, MAX_AVATAR_BYTES);
        if (!bytes) {
          return reply.code(413).send({ error: 'avatar-too-large' });
        }
      }
    }
    if (!bytes || !mime) {
      return reply.code(400).send({ error: 'missing-avatar' });
    }
    if (!ALLOWED_AVATAR_MIMES.has(mime)) {
      return reply.code(415).send({ error: 'unsupported-mime' });
    }
    store.setAvatar(user.id, mime, new Uint8Array(bytes));
    return reply.code(204).send();
  });

  // ── DELETE /auth/profile/avatar ─────────────────────────────────────
  app.delete('/auth/profile/avatar', async (req, reply) => {
    const user = requireUser(req, reply, store);
    if (!user) return;
    store.setAvatar(user.id, null, null);
    return reply.code(204).send();
  });

  // ── GET /auth/profile/avatar (serve bytes) ──────────────────────────
  // Public-ish — any signed-in user can fetch any other user's
  // avatar via `?id=<userId>`. Same trust model as a display name.
  // Team mode will scope this to workspace members.
  app.get<{ Querystring: { id?: string } }>('/auth/profile/avatar', async (req, reply) => {
    const me = requireUser(req, reply, store);
    if (!me) return;
    const targetId = req.query.id ? Number(req.query.id) : me.id;
    if (!Number.isFinite(targetId) || targetId <= 0) {
      return reply.code(400).send({ error: 'bad-id' });
    }
    const avatar = store.getAvatar(targetId);
    if (!avatar) return reply.code(404).send({ error: 'no-avatar' });
    reply
      .header('content-type', avatar.mime)
      .header('content-length', String(avatar.bytes.byteLength))
      // Short cache — clients re-fetch every couple of minutes so
      // an avatar change shows up reasonably quickly across tabs.
      .header('cache-control', 'private, max-age=120');
    return reply.send(Buffer.from(avatar.bytes));
  });
}

function requireUser(
  req: FastifyRequest,
  reply: FastifyReply,
  store: PersonalAuthStore,
): PublicUser | null {
  if (store.mode === 'none') {
    reply.code(503).send({ error: 'personal-mode-disabled' });
    return null;
  }
  const user = currentUser(req, store);
  if (!user) {
    reply.code(401).send({ error: 'unauthenticated' });
    return null;
  }
  return user;
}

async function streamToBuffer(
  stream: NodeJS.ReadableStream | Readable,
  maxBytes: number,
): Promise<Buffer | null> {
  const chunks: Buffer[] = [];
  let total = 0;
  for await (const chunk of stream) {
    let buf: Buffer;
    if (typeof chunk === 'string') buf = Buffer.from(chunk);
    else if (Buffer.isBuffer(chunk)) buf = chunk;
    else {
      const v = chunk as ArrayBufferView;
      buf = Buffer.from(v.buffer, v.byteOffset, v.byteLength);
    }
    total += buf.length;
    if (total > maxBytes) return null;
    chunks.push(buf);
  }
  return Buffer.concat(chunks);
}

// Re-export so file-actions.ts can typecheck without a circular import.
export type { UserProfile };
