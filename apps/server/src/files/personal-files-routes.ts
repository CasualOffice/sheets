import type { FastifyInstance, FastifyReply, FastifyRequest } from 'fastify';
import { Readable } from 'node:stream';
import type { PersonalAuthStore, PublicUser } from '../auth/personal.js';
import { currentUser } from '../auth/personal-routes.js';
import type { HostIntegration } from '../host/integration.js';
import { newVersion } from '../host/integration.js';

/**
 * Personal-mode `/files/*` routes (Phase C Batch 2 of #49).
 *
 * The byte payload still lives in the configured `HostIntegration`
 * backend (local / s3 / postgres). This module adds a SQLite-backed
 * ownership + listing layer on top — the registry rows in
 * `personal.files`. Every route enforces ownership at the registry
 * level before touching the host integration.
 *
 * Authoring this as a separate route surface from the existing WOPI
 * endpoints keeps each model honest:
 *
 *   - `/wopi/files/:id` — JWT-secured (or anonymous-by-URL when the
 *     secret is unset). Used by embedded WOPI hosts, the admin
 *     panel, and the in-browser collab seed flow. No cookie auth.
 *   - `/files/*`        — cookie-auth (cs_session). User-scoped:
 *     listing + ownership checks are done against the personal
 *     `files` table.
 *
 *   The two surfaces never accidentally trade scopes: a session
 *   cookie can't unlock `/wopi/files/:id`, and a JWT can't unlock
 *   `/files/*`.
 */

const MAX_DISPLAY_NAME = 255;

export function registerPersonalFilesRoutes(
  app: FastifyInstance,
  store: PersonalAuthStore,
  host: HostIntegration,
  opts: { maxUploadBytes: number },
): void {
  // ── GET /files ──────────────────────────────────────────────────────
  app.get('/files', async (req, reply) => {
    const user = requireUser(req, reply, store);
    if (!user) return;
    const files = store.listFilesForUser(user.id);
    return reply.send({
      files: files.map((f) => ({
        id: f.id,
        name: f.displayName,
        size: f.size,
        etag: f.etag,
        createdAt: f.createdAt,
        modifiedAt: f.modifiedAt,
      })),
    });
  });

  // ── POST /files ─────────────────────────────────────────────────────
  // Multipart upload. The file part is required; an optional `name`
  // text part overrides the filename. Browser sends the .xlsx as
  // `file` per the standard form pattern.
  app.post('/files', async (req, reply) => {
    const user = requireUser(req, reply, store);
    if (!user) return;
    const parts = req.parts();
    let bytes: Buffer | null = null;
    let baseName: string | null = null;

    for await (const part of parts) {
      if (part.type === 'file' && part.fieldname === 'file') {
        baseName = sanitiseFilename(part.filename);
        bytes = await streamToBuffer(part.file, opts.maxUploadBytes);
        if (!bytes) {
          return reply.code(413).send({ error: 'upload-too-large' });
        }
      } else if (part.type === 'field' && part.fieldname === 'name') {
        const v = String(part.value ?? '').trim();
        if (v) baseName = sanitiseFilename(v);
      }
    }

    if (!bytes || !baseName) {
      return reply.code(400).send({ error: 'missing-file' });
    }

    const etag = newVersion();
    const record = store.createFile({
      ownerId: user.id,
      displayName: baseName,
      size: bytes.length,
      etag,
    });
    try {
      await host.putFile(record.id, new Uint8Array(bytes), {
        fileName: baseName,
      });
    } catch (err) {
      // Roll back the registry row so a backend hiccup doesn't leave
      // a stale entry the user sees in their list but can't open.
      store.deleteFile(record.id);
      throw err;
    }
    return reply.code(201).send({
      file: {
        id: record.id,
        name: record.displayName,
        size: record.size,
        etag: record.etag,
        createdAt: record.createdAt,
        modifiedAt: record.modifiedAt,
      },
    });
  });

  // ── GET /files/:id (download bytes) ─────────────────────────────────
  app.get<{ Params: { id: string } }>('/files/:id', async (req, reply) => {
    const user = requireUser(req, reply, store);
    if (!user) return;
    const record = ownedFileOr403(store, req.params.id, user, reply);
    if (!record) return;
    const bytes = await host.getFile(record.id);
    if (!bytes) return reply.code(404).send({ error: 'not-found' });
    reply
      .header('content-type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
      .header('etag', record.etag)
      .header('content-length', String(bytes.byteLength))
      .header(
        'content-disposition',
        `attachment; filename="${encodeURIComponent(record.displayName)}"`,
      );
    return reply.send(Buffer.from(bytes));
  });

  // ── POST /files/:id (replace bytes) ─────────────────────────────────
  // Body is the raw xlsx bytes (application/octet-stream). If-Match
  // header carries the etag the client started from; mismatch returns
  // 412 so the web app can show the conflict modal.
  app.post<{ Params: { id: string } }>('/files/:id', async (req, reply) => {
    const user = requireUser(req, reply, store);
    if (!user) return;
    const record = ownedFileOr403(store, req.params.id, user, reply);
    if (!record) return;
    const ifMatch = (req.headers['if-match'] as string | undefined)?.trim();
    if (ifMatch && ifMatch !== record.etag) {
      return reply.code(412).send({
        error: 'version-mismatch',
        expected: record.etag,
      });
    }
    const body = req.body;
    const bytes =
      body instanceof Uint8Array ? Buffer.from(body) : Buffer.isBuffer(body) ? body : null;
    if (!bytes) return reply.code(400).send({ error: 'expected-binary-body' });
    if (bytes.length > opts.maxUploadBytes) {
      return reply.code(413).send({ error: 'upload-too-large' });
    }

    // Conflict detection lives in this layer (against the registry
    // row), not in the host integration. We deliberately do not
    // forward `ifMatchVersion` to `putFile` — the host's internal
    // version sequence is its own bookkeeping and would only get in
    // the way here, since `/files/*` clients see and reason about
    // the registry etag.
    const newEtag = newVersion();
    await host.putFile(record.id, new Uint8Array(bytes), {
      fileName: record.displayName,
    });
    store.recordFileUpdate(record.id, bytes.length, newEtag);
    return reply.send({
      file: {
        id: record.id,
        name: record.displayName,
        size: bytes.length,
        etag: newEtag,
        createdAt: record.createdAt,
        modifiedAt: Date.now(),
      },
    });
  });

  // ── DELETE /files/:id ───────────────────────────────────────────────
  app.delete<{ Params: { id: string } }>('/files/:id', async (req, reply) => {
    const user = requireUser(req, reply, store);
    if (!user) return;
    const record = ownedFileOr403(store, req.params.id, user, reply);
    if (!record) return;
    store.deleteFile(record.id);
    if (host.deleteFile) await host.deleteFile(record.id);
    return reply.code(204).send();
  });

  // ── PATCH /files/:id (rename) ───────────────────────────────────────
  app.patch<{ Params: { id: string }; Body: { name?: unknown } }>(
    '/files/:id',
    async (req, reply) => {
      const user = requireUser(req, reply, store);
      if (!user) return;
      const record = ownedFileOr403(store, req.params.id, user, reply);
      if (!record) return;
      const body = (req.body ?? {}) as { name?: unknown };
      if (typeof body.name !== 'string') {
        return reply.code(400).send({ error: 'missing-name' });
      }
      const newName = sanitiseFilename(body.name);
      if (!newName) return reply.code(400).send({ error: 'invalid-name' });
      const ok = store.renameFile(record.id, newName);
      if (!ok) return reply.code(404).send({ error: 'not-found' });
      return reply.code(204).send();
    },
  );
}

// ── Helpers ──────────────────────────────────────────────────────────

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

function ownedFileOr403(
  store: PersonalAuthStore,
  id: string,
  user: PublicUser,
  reply: FastifyReply,
) {
  const record = store.getFile(id);
  if (!record) {
    reply.code(404).send({ error: 'not-found' });
    return null;
  }
  if (record.ownerId !== user.id) {
    // 404 not 403 — don't leak whether the file exists to non-owners.
    // The admin role intentionally has no cross-user visibility here
    // (per the Phase C decision in docs/STORAGE_MODES.md).
    reply.code(404).send({ error: 'not-found' });
    return null;
  }
  return record;
}

function sanitiseFilename(raw: string): string {
  // Strip directory separators + control chars; cap length. A real
  // filename has at most 255 chars on most filesystems.
  const cleaned = raw
    .replace(/[\\/]/g, '_')
    // eslint-disable-next-line no-control-regex
    .replace(/[\x00-\x1f]/g, '')
    .trim()
    .slice(0, MAX_DISPLAY_NAME);
  return cleaned;
}

async function streamToBuffer(
  stream: NodeJS.ReadableStream | Readable,
  maxBytes: number,
): Promise<Buffer | null> {
  const chunks: Buffer[] = [];
  let total = 0;
  for await (const chunk of stream) {
    let buf: Buffer;
    if (typeof chunk === 'string') {
      buf = Buffer.from(chunk);
    } else if (Buffer.isBuffer(chunk)) {
      buf = chunk;
    } else {
      // ArrayBufferView from a Web stream; copy into a Buffer via the
      // underlying byte range. Using new Uint8Array(view.buffer, view.byteOffset, view.byteLength)
      // avoids the "ArrayBuffer not assignable" overload mismatch on Node 22.
      const view = chunk as ArrayBufferView;
      buf = Buffer.from(view.buffer, view.byteOffset, view.byteLength);
    }
    total += buf.length;
    if (total > maxBytes) return null;
    chunks.push(buf);
  }
  return Buffer.concat(chunks);
}
