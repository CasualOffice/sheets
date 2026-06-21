/**
 * Integration test for the `/files/:id/shares*` routes (sharing-model
 * §6.1, the SAFE FOUNDATION). Boots a Fastify app in-process via
 * `inject`, wires the personal auth store + an in-memory host, and
 * walks the link CRUD lifecycle plus the ownership gate.
 *
 *   create → list → patch → delete, password omission from list,
 *   role validation (400), expiry validation (400), non-owner 404,
 *   admin cross-user reach, anonymous 401, mode=none 503.
 *
 * These routes are inert — minting a token grants no access until the
 * join-handshake batch wires enforcement. This spec only asserts the
 * persistence + gating contract.
 *
 * Run with `pnpm test:unit`.
 */
import { strict as assert } from 'node:assert';
import { test } from 'node:test';
import { mkdtempSync, rmSync } from 'node:fs';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import Fastify, { type FastifyInstance } from 'fastify';
import multipart from '@fastify/multipart';
import cookie from '@fastify/cookie';

import { PersonalAuthStore } from '../auth/personal.js';
import { registerPersonalAuthRoutes } from '../auth/personal-routes.js';
import { registerPersonalFilesRoutes } from './personal-files-routes.js';
import { registerPersonalSharesRoutes } from './personal-shares-routes.js';
import { MemoryHost } from '../host/memory.js';

async function makeApp(opts: { mode: 'single' | 'multi' | 'none' } = { mode: 'multi' }) {
  const dir = mkdtempSync(join(tmpdir(), 'casual-shares-routes-'));
  const store = new PersonalAuthStore({
    dbPath: join(dir, 'users.db'),
    mode: opts.mode,
    bootstrap: null,
  });
  const host = new MemoryHost();
  const app = Fastify({ logger: false, bodyLimit: 5 * 1024 * 1024 });
  await app.register(cookie);
  await app.register(multipart, { limits: { fileSize: 5 * 1024 * 1024 } });
  registerPersonalAuthRoutes(app, store);
  registerPersonalFilesRoutes(app, store, host, { maxUploadBytes: 5 * 1024 * 1024 });
  // A fixed set of "live" rooms so the share-link route's roomId
  // existence check passes for the canonical test room and rejects an
  // unknown one.
  const liveRooms = new Set(['room-1', 'room-2']);
  registerPersonalSharesRoutes(app, store, { roomExists: (id) => liveRooms.has(id) });
  await app.ready();
  return {
    app,
    store,
    cleanup: async () => {
      await app.close();
      store.close();
      rmSync(dir, { recursive: true, force: true });
    },
  };
}

async function signup(app: FastifyInstance, username: string, password: string): Promise<string> {
  const r = await app.inject({
    method: 'POST',
    url: '/auth/signup',
    payload: { username, password },
  });
  assert.equal(r.statusCode, 201, `signup expected 201, got ${r.statusCode}: ${r.body}`);
  const setCookie = r.cookies.find((c) => c.name === 'cs_session');
  assert.ok(setCookie, 'expected cs_session cookie');
  return `cs_session=${setCookie.value}`;
}

function multipartBody(content: Buffer, filename: string, boundary: string): Buffer {
  const head = Buffer.from(
    `--${boundary}\r\n` +
      `content-disposition: form-data; name="file"; filename="${filename}"\r\n` +
      `content-type: application/octet-stream\r\n\r\n`,
    'utf8',
  );
  const tail = Buffer.from(`\r\n--${boundary}--\r\n`, 'utf8');
  return Buffer.concat([head, content, tail]);
}

async function uploadFile(
  app: FastifyInstance,
  cookie: string,
  name = 'book.xlsx',
): Promise<string> {
  const boundary = '----CasualTestBoundary' + Math.random().toString(36).slice(2);
  const body = multipartBody(Buffer.from('FAKE-XLSX'), name, boundary);
  const r = await app.inject({
    method: 'POST',
    url: '/files',
    headers: {
      cookie,
      'content-type': `multipart/form-data; boundary=${boundary}`,
      'content-length': String(body.length),
    },
    payload: body,
  });
  assert.equal(r.statusCode, 201, `upload status ${r.statusCode}: ${r.body}`);
  return (JSON.parse(r.body).file as { id: string }).id;
}

test('share link CRUD: create → list → patch → delete', async () => {
  const { app, cleanup } = await makeApp();
  try {
    const cookie = await signup(app, 'alice', 'longpassword');
    const fileId = await uploadFile(app, cookie);

    // Empty list to start.
    let r = await app.inject({
      method: 'GET',
      url: `/files/${fileId}/shares`,
      headers: { cookie },
    });
    assert.equal(r.statusCode, 200);
    assert.deepEqual(JSON.parse(r.body), { links: [] });

    // Mint a link with a password + expiry.
    r = await app.inject({
      method: 'POST',
      url: `/files/${fileId}/shares/link`,
      headers: { cookie, 'content-type': 'application/json' },
      payload: { roomId: 'room-1', role: 'edit', expiresInDays: 7, password: 'secret' },
    });
    assert.equal(r.statusCode, 201, r.body);
    const minted = JSON.parse(r.body) as {
      token: string;
      roomId: string;
      role: string;
      expiresAt: number;
      url: string;
    };
    assert.equal(minted.role, 'edit');
    assert.equal(minted.roomId, 'room-1', 'mint echoes the bound roomId');
    assert.ok(minted.token.length >= 40);
    assert.ok(minted.expiresAt > Date.now());
    assert.equal(minted.url, `?share=${minted.token}`);

    // List shows it — passwordHash is never present, hasPassword is.
    r = await app.inject({ method: 'GET', url: `/files/${fileId}/shares`, headers: { cookie } });
    const links = JSON.parse(r.body).links as Array<Record<string, unknown>>;
    assert.equal(links.length, 1);
    assert.equal(links[0]?.hasPassword, true);
    assert.equal(links[0]?.passwordHash, undefined, 'passwordHash must be omitted');
    assert.ok(!('passwordHash' in (links[0] as object)));

    // Patch role + clear expiry.
    r = await app.inject({
      method: 'PATCH',
      url: `/files/${fileId}/shares/link/${minted.token}`,
      headers: { cookie, 'content-type': 'application/json' },
      payload: { role: 'view' },
    });
    assert.equal(r.statusCode, 200, r.body);
    assert.equal(JSON.parse(r.body).role, 'view');

    // Delete.
    r = await app.inject({
      method: 'DELETE',
      url: `/files/${fileId}/shares/link/${minted.token}`,
      headers: { cookie },
    });
    assert.equal(r.statusCode, 204);
    r = await app.inject({ method: 'GET', url: `/files/${fileId}/shares`, headers: { cookie } });
    assert.deepEqual(JSON.parse(r.body), { links: [] });
  } finally {
    await cleanup();
  }
});

test('share link: a no-password link reports hasPassword=false', async () => {
  const { app, cleanup } = await makeApp();
  try {
    const cookie = await signup(app, 'alice', 'longpassword');
    const fileId = await uploadFile(app, cookie);
    const r = await app.inject({
      method: 'POST',
      url: `/files/${fileId}/shares/link`,
      headers: { cookie, 'content-type': 'application/json' },
      payload: { roomId: 'room-1', role: 'view' },
    });
    assert.equal(r.statusCode, 201);
    assert.equal(JSON.parse(r.body).expiresAt, null);
    const list = await app.inject({
      method: 'GET',
      url: `/files/${fileId}/shares`,
      headers: { cookie },
    });
    assert.equal(
      (JSON.parse(list.body).links as Array<{ hasPassword: boolean }>)[0]?.hasPassword,
      false,
    );
  } finally {
    await cleanup();
  }
});

test('share link: role validation rejects bad roles with 400', async () => {
  const { app, cleanup } = await makeApp();
  try {
    const cookie = await signup(app, 'alice', 'longpassword');
    const fileId = await uploadFile(app, cookie);
    for (const role of ['admin', '', 123, undefined]) {
      const r = await app.inject({
        method: 'POST',
        url: `/files/${fileId}/shares/link`,
        headers: { cookie, 'content-type': 'application/json' },
        payload: { roomId: 'room-1', role },
      });
      assert.equal(r.statusCode, 400, `role=${String(role)} should 400`);
      assert.equal(JSON.parse(r.body).error, 'invalid-role');
    }
  } finally {
    await cleanup();
  }
});

test('share link: expiry validation rejects non-positive values with 400', async () => {
  const { app, cleanup } = await makeApp();
  try {
    const cookie = await signup(app, 'alice', 'longpassword');
    const fileId = await uploadFile(app, cookie);
    // Note: NaN/Infinity can't survive JSON transport (they serialise
    // to null), so the over-the-wire bad values are 0, negative, and
    // non-numeric. The numeric edge cases are covered store-side.
    for (const expiresInDays of [0, -1, 'soon']) {
      const r = await app.inject({
        method: 'POST',
        url: `/files/${fileId}/shares/link`,
        headers: { cookie, 'content-type': 'application/json' },
        payload: { roomId: 'room-1', role: 'view', expiresInDays },
      });
      assert.equal(r.statusCode, 400, `expiresInDays=${String(expiresInDays)} should 400`);
      assert.equal(JSON.parse(r.body).error, 'invalid-expiry');
    }
  } finally {
    await cleanup();
  }
});

test('share link: non-owner gets 404 (no existence leak), owner unaffected', async () => {
  const { app, cleanup } = await makeApp({ mode: 'multi' });
  try {
    const aliceCookie = await signup(app, 'alice', 'longpassword');
    const bobCookie = await signup(app, 'bob', 'longpassword');
    const fileId = await uploadFile(app, aliceCookie);

    // Bob (non-admin member) can't see, mint, patch, or delete.
    let r = await app.inject({
      method: 'GET',
      url: `/files/${fileId}/shares`,
      headers: { cookie: bobCookie },
    });
    assert.equal(r.statusCode, 404);
    r = await app.inject({
      method: 'POST',
      url: `/files/${fileId}/shares/link`,
      headers: { cookie: bobCookie, 'content-type': 'application/json' },
      payload: { roomId: 'room-1', role: 'view' },
    });
    assert.equal(r.statusCode, 404);

    // Alice still works.
    r = await app.inject({
      method: 'POST',
      url: `/files/${fileId}/shares/link`,
      headers: { cookie: aliceCookie, 'content-type': 'application/json' },
      payload: { roomId: 'room-1', role: 'view' },
    });
    assert.equal(r.statusCode, 201);
  } finally {
    await cleanup();
  }
});

test('share link: admin reaches another user file (§4 RequireAdmin)', async () => {
  const { app, cleanup } = await makeApp({ mode: 'multi' });
  try {
    // First account is admin.
    const adminCookie = await signup(app, 'admin', 'longpassword');
    const bobCookie = await signup(app, 'bob', 'longpassword');
    const bobFile = await uploadFile(app, bobCookie);

    const r = await app.inject({
      method: 'POST',
      url: `/files/${bobFile}/shares/link`,
      headers: { cookie: adminCookie, 'content-type': 'application/json' },
      payload: { roomId: 'room-1', role: 'edit' },
    });
    assert.equal(r.statusCode, 201, 'admin should reach any file');
  } finally {
    await cleanup();
  }
});

test('share link: token from a different workbook returns 404 on patch/delete', async () => {
  const { app, cleanup } = await makeApp();
  try {
    const cookie = await signup(app, 'alice', 'longpassword');
    const fileA = await uploadFile(app, cookie, 'a.xlsx');
    const fileB = await uploadFile(app, cookie, 'b.xlsx');

    const mint = await app.inject({
      method: 'POST',
      url: `/files/${fileA}/shares/link`,
      headers: { cookie, 'content-type': 'application/json' },
      payload: { roomId: 'room-1', role: 'view' },
    });
    const token = JSON.parse(mint.body).token as string;

    // Same owner, but the token belongs to fileA — editing it via
    // fileB's path must 404.
    const r = await app.inject({
      method: 'DELETE',
      url: `/files/${fileB}/shares/link/${token}`,
      headers: { cookie },
    });
    assert.equal(r.statusCode, 404);
  } finally {
    await cleanup();
  }
});

test('share link: missing / non-string roomId rejected with 400 (invalid-room)', async () => {
  const { app, cleanup } = await makeApp();
  try {
    const cookie = await signup(app, 'alice', 'longpassword');
    const fileId = await uploadFile(app, cookie);
    // roomId omitted entirely, and a few non-string shapes.
    for (const roomId of [undefined, 123, '', '   ']) {
      const payload: Record<string, unknown> = { role: 'view' };
      if (roomId !== undefined) payload.roomId = roomId;
      const r = await app.inject({
        method: 'POST',
        url: `/files/${fileId}/shares/link`,
        headers: { cookie, 'content-type': 'application/json' },
        payload,
      });
      assert.equal(r.statusCode, 400, `roomId=${String(roomId)} should 400`);
      assert.equal(JSON.parse(r.body).error, 'invalid-room');
    }
  } finally {
    await cleanup();
  }
});

test('share link: unknown roomId rejected with 404 (room-not-found)', async () => {
  const { app, cleanup } = await makeApp();
  try {
    const cookie = await signup(app, 'alice', 'longpassword');
    const fileId = await uploadFile(app, cookie);
    const r = await app.inject({
      method: 'POST',
      url: `/files/${fileId}/shares/link`,
      headers: { cookie, 'content-type': 'application/json' },
      payload: { roomId: 'room-does-not-exist', role: 'view' },
    });
    assert.equal(r.statusCode, 404);
    assert.equal(JSON.parse(r.body).error, 'room-not-found');
  } finally {
    await cleanup();
  }
});

test('share link: list + mint surface the bound roomId', async () => {
  const { app, cleanup } = await makeApp();
  try {
    const cookie = await signup(app, 'alice', 'longpassword');
    const fileId = await uploadFile(app, cookie);
    const mint = await app.inject({
      method: 'POST',
      url: `/files/${fileId}/shares/link`,
      headers: { cookie, 'content-type': 'application/json' },
      payload: { roomId: 'room-2', role: 'edit' },
    });
    assert.equal(mint.statusCode, 201);
    assert.equal(JSON.parse(mint.body).roomId, 'room-2');

    const list = await app.inject({
      method: 'GET',
      url: `/files/${fileId}/shares`,
      headers: { cookie },
    });
    const links = JSON.parse(list.body).links as Array<{ roomId: string }>;
    assert.equal(links[0]?.roomId, 'room-2', 'list surfaces roomId');
  } finally {
    await cleanup();
  }
});

test('share link: anonymous → 401 on every route', async () => {
  const { app, cleanup } = await makeApp();
  try {
    for (const route of [
      ['GET', '/files/x/shares'],
      ['POST', '/files/x/shares/link'],
      ['PATCH', '/files/x/shares/link/t'],
      ['DELETE', '/files/x/shares/link/t'],
    ] as const) {
      const r = await app.inject({ method: route[0], url: route[1] });
      assert.equal(r.statusCode, 401, `${route[0]} ${route[1]} should 401, got ${r.statusCode}`);
    }
  } finally {
    await cleanup();
  }
});

test("share link: mode 'none' shadows the routes with 503", async () => {
  const { app, cleanup } = await makeApp({ mode: 'none' });
  try {
    const r = await app.inject({ method: 'GET', url: '/files/x/shares' });
    assert.equal(r.statusCode, 503);
  } finally {
    await cleanup();
  }
});

// ── Public /meta endpoint (sharing-model §6.1 — pre-join discovery) ─────

test('share link /meta: valid token returns role + hasPassword + roomId (no hash)', async () => {
  const { app, cleanup } = await makeApp();
  try {
    const cookie = await signup(app, 'alice', 'longpassword');
    const fileId = await uploadFile(app, cookie);
    const mint = await app.inject({
      method: 'POST',
      url: `/files/${fileId}/shares/link`,
      headers: { cookie, 'content-type': 'application/json' },
      payload: { roomId: 'room-1', role: 'edit', password: 'secret' },
    });
    assert.equal(mint.statusCode, 201, mint.body);
    const token = JSON.parse(mint.body).token as string;

    // No auth header — the token IS the capability, so /meta is public.
    const r = await app.inject({ method: 'GET', url: `/files/shares/link/${token}/meta` });
    assert.equal(r.statusCode, 200, r.body);
    const body = JSON.parse(r.body) as Record<string, unknown>;
    assert.equal(body.valid, true);
    assert.equal(body.role, 'edit');
    assert.equal(body.hasPassword, true);
    assert.equal(body.roomId, 'room-1');
    // The bcrypt hash must NEVER appear in the public response.
    assert.equal(body.passwordHash, undefined, 'passwordHash must not leak');
    assert.ok(!('passwordHash' in body), 'no passwordHash key at all');
  } finally {
    await cleanup();
  }
});

test('share link /meta: a no-password link reports hasPassword=false', async () => {
  const { app, cleanup } = await makeApp();
  try {
    const cookie = await signup(app, 'alice', 'longpassword');
    const fileId = await uploadFile(app, cookie);
    const mint = await app.inject({
      method: 'POST',
      url: `/files/${fileId}/shares/link`,
      headers: { cookie, 'content-type': 'application/json' },
      payload: { roomId: 'room-2', role: 'view' },
    });
    const token = JSON.parse(mint.body).token as string;

    const r = await app.inject({ method: 'GET', url: `/files/shares/link/${token}/meta` });
    assert.equal(r.statusCode, 200);
    const body = JSON.parse(r.body) as Record<string, unknown>;
    assert.equal(body.valid, true);
    assert.equal(body.role, 'view');
    assert.equal(body.hasPassword, false);
    assert.equal(body.roomId, 'room-2');
  } finally {
    await cleanup();
  }
});

test('share link /meta: unknown token returns { valid: false } and nothing else', async () => {
  const { app, cleanup } = await makeApp();
  try {
    const r = await app.inject({
      method: 'GET',
      url: '/files/shares/link/this-token-does-not-exist/meta',
    });
    assert.equal(r.statusCode, 200);
    assert.deepEqual(JSON.parse(r.body), { valid: false });
  } finally {
    await cleanup();
  }
});

test('share link /meta: expired token returns { valid: false } (no leak of role/room)', async () => {
  const { app, store, cleanup } = await makeApp();
  try {
    const cookie = await signup(app, 'alice', 'longpassword');
    const fileId = await uploadFile(app, cookie);
    // Mint directly through the store so we can backdate the expiry —
    // the route caps expiresInDays at a positive value, so an
    // already-lapsed link can only be made store-side.
    const link = store.createShareLink({
      workbookId: fileId,
      roomId: 'room-1',
      role: 'edit',
      createdBy: 1,
      expiresAt: Date.now() - 1000,
      password: 'secret',
    });
    const r = await app.inject({
      method: 'GET',
      url: `/files/shares/link/${link.token}/meta`,
    });
    assert.equal(r.statusCode, 200);
    // Expired collapses to the same shape as unknown — a probe can't
    // learn the token ever existed (or its role/room/hasPassword).
    assert.deepEqual(JSON.parse(r.body), { valid: false });
  } finally {
    await cleanup();
  }
});

test("share link /meta: mode 'none' → 503 (feature off, distinguishable from bad token)", async () => {
  const { app, cleanup } = await makeApp({ mode: 'none' });
  try {
    const r = await app.inject({ method: 'GET', url: '/files/shares/link/anything/meta' });
    assert.equal(r.statusCode, 503);
  } finally {
    await cleanup();
  }
});

// ── Member ACL routes (sharing-model §6.2 — MULTI MODE ONLY) ────────────

test('member ACL: add → list → patch → delete by username handle', async () => {
  const { app, store, cleanup } = await makeApp({ mode: 'multi' });
  try {
    const aliceCookie = await signup(app, 'alice', 'longpassword');
    await signup(app, 'bob', 'longpassword');
    const fileId = await uploadFile(app, aliceCookie);

    // Empty to start.
    let r = await app.inject({
      method: 'GET',
      url: `/files/${fileId}/shares/members`,
      headers: { cookie: aliceCookie },
    });
    assert.equal(r.statusCode, 200);
    assert.deepEqual(JSON.parse(r.body), { members: [] });

    // Add bob by username.
    r = await app.inject({
      method: 'POST',
      url: `/files/${fileId}/shares/member`,
      headers: { cookie: aliceCookie, 'content-type': 'application/json' },
      payload: { handle: 'bob', role: 'edit' },
    });
    assert.equal(r.statusCode, 201, r.body);
    const added = JSON.parse(r.body) as { memberId: number; username: string; role: string };
    assert.equal(added.username, 'bob');
    assert.equal(added.role, 'edit');
    const bobId = added.memberId;

    // List surfaces it with username + role.
    r = await app.inject({
      method: 'GET',
      url: `/files/${fileId}/shares/members`,
      headers: { cookie: aliceCookie },
    });
    const members = JSON.parse(r.body).members as Array<Record<string, unknown>>;
    assert.equal(members.length, 1);
    assert.equal(members[0]?.username, 'bob');
    assert.equal(members[0]?.role, 'edit');

    // Patch role.
    r = await app.inject({
      method: 'PATCH',
      url: `/files/${fileId}/shares/member/${bobId}`,
      headers: { cookie: aliceCookie, 'content-type': 'application/json' },
      payload: { role: 'view' },
    });
    assert.equal(r.statusCode, 200, r.body);
    assert.equal(JSON.parse(r.body).role, 'view');
    assert.equal(store.getMemberRole(fileId, bobId), 'view');

    // Delete.
    r = await app.inject({
      method: 'DELETE',
      url: `/files/${fileId}/shares/member/${bobId}`,
      headers: { cookie: aliceCookie },
    });
    assert.equal(r.statusCode, 204);
    assert.equal(store.getMemberRole(fileId, bobId), null);
  } finally {
    await cleanup();
  }
});

test('member ACL: add resolves an email handle', async () => {
  const { app, store, cleanup } = await makeApp({ mode: 'multi' });
  try {
    const aliceCookie = await signup(app, 'alice', 'longpassword');
    await signup(app, 'bob', 'longpassword');
    const bobId = store.findMemberByHandle('bob')!.id;
    store.updateProfile(bobId, { email: 'bob@example.com' });
    const fileId = await uploadFile(app, aliceCookie);

    const r = await app.inject({
      method: 'POST',
      url: `/files/${fileId}/shares/member`,
      headers: { cookie: aliceCookie, 'content-type': 'application/json' },
      payload: { handle: 'BOB@EXAMPLE.COM', role: 'view' },
    });
    assert.equal(r.statusCode, 201, r.body);
    assert.equal(JSON.parse(r.body).memberId, bobId);
    assert.equal(JSON.parse(r.body).email, 'bob@example.com');
  } finally {
    await cleanup();
  }
});

test('member ACL: unknown handle → 404 user-not-found', async () => {
  const { app, cleanup } = await makeApp({ mode: 'multi' });
  try {
    const aliceCookie = await signup(app, 'alice', 'longpassword');
    const fileId = await uploadFile(app, aliceCookie);
    const r = await app.inject({
      method: 'POST',
      url: `/files/${fileId}/shares/member`,
      headers: { cookie: aliceCookie, 'content-type': 'application/json' },
      payload: { handle: 'ghost', role: 'view' },
    });
    assert.equal(r.statusCode, 404);
    assert.equal(JSON.parse(r.body).error, 'user-not-found');
  } finally {
    await cleanup();
  }
});

test('member ACL: cannot add yourself / the owner', async () => {
  const { app, cleanup } = await makeApp({ mode: 'multi' });
  try {
    const aliceCookie = await signup(app, 'alice', 'longpassword');
    const fileId = await uploadFile(app, aliceCookie);
    const r = await app.inject({
      method: 'POST',
      url: `/files/${fileId}/shares/member`,
      headers: { cookie: aliceCookie, 'content-type': 'application/json' },
      payload: { handle: 'alice', role: 'view' },
    });
    assert.equal(r.statusCode, 400);
    assert.equal(JSON.parse(r.body).error, 'cannot-add-owner');
  } finally {
    await cleanup();
  }
});

test('member ACL: role + handle validation (400)', async () => {
  const { app, cleanup } = await makeApp({ mode: 'multi' });
  try {
    const aliceCookie = await signup(app, 'alice', 'longpassword');
    await signup(app, 'bob', 'longpassword');
    const fileId = await uploadFile(app, aliceCookie);

    // Bad role.
    let r = await app.inject({
      method: 'POST',
      url: `/files/${fileId}/shares/member`,
      headers: { cookie: aliceCookie, 'content-type': 'application/json' },
      payload: { handle: 'bob', role: 'admin' },
    });
    assert.equal(r.statusCode, 400);
    assert.equal(JSON.parse(r.body).error, 'invalid-role');

    // Bad handle.
    for (const handle of [undefined, '', '   ', 123]) {
      const payload: Record<string, unknown> = { role: 'view' };
      if (handle !== undefined) payload.handle = handle;
      r = await app.inject({
        method: 'POST',
        url: `/files/${fileId}/shares/member`,
        headers: { cookie: aliceCookie, 'content-type': 'application/json' },
        payload,
      });
      assert.equal(r.statusCode, 400, `handle=${String(handle)} should 400`);
      assert.equal(JSON.parse(r.body).error, 'invalid-handle');
    }
  } finally {
    await cleanup();
  }
});

test('member ACL: patch/delete on a non-existent ACL → 404', async () => {
  const { app, store, cleanup } = await makeApp({ mode: 'multi' });
  try {
    const aliceCookie = await signup(app, 'alice', 'longpassword');
    await signup(app, 'bob', 'longpassword');
    const bobId = store.findMemberByHandle('bob')!.id;
    const fileId = await uploadFile(app, aliceCookie);

    let r = await app.inject({
      method: 'PATCH',
      url: `/files/${fileId}/shares/member/${bobId}`,
      headers: { cookie: aliceCookie, 'content-type': 'application/json' },
      payload: { role: 'view' },
    });
    assert.equal(r.statusCode, 404);

    r = await app.inject({
      method: 'DELETE',
      url: `/files/${fileId}/shares/member/${bobId}`,
      headers: { cookie: aliceCookie },
    });
    assert.equal(r.statusCode, 404);
  } finally {
    await cleanup();
  }
});

test('member ACL: invalid :memberId path param → 400', async () => {
  const { app, cleanup } = await makeApp({ mode: 'multi' });
  try {
    const aliceCookie = await signup(app, 'alice', 'longpassword');
    const fileId = await uploadFile(app, aliceCookie);
    const r = await app.inject({
      method: 'DELETE',
      url: `/files/${fileId}/shares/member/not-a-number`,
      headers: { cookie: aliceCookie },
    });
    assert.equal(r.statusCode, 400);
    assert.equal(JSON.parse(r.body).error, 'invalid-member');
  } finally {
    await cleanup();
  }
});

test('member ACL: non-owner gets 404, admin reaches any file', async () => {
  const { app, cleanup } = await makeApp({ mode: 'multi' });
  try {
    // First account is admin.
    const adminCookie = await signup(app, 'admin', 'longpassword');
    const bobCookie = await signup(app, 'bob', 'longpassword');
    await signup(app, 'carol', 'longpassword');
    const bobFile = await uploadFile(app, bobCookie);

    // Carol (non-owner, non-admin) → 404, no existence leak.
    const carolCookie = await signup(app, 'dave', 'longpassword');
    let r = await app.inject({
      method: 'GET',
      url: `/files/${bobFile}/shares/members`,
      headers: { cookie: carolCookie },
    });
    assert.equal(r.statusCode, 404);

    // Admin reaches bob's file and can add carol.
    r = await app.inject({
      method: 'POST',
      url: `/files/${bobFile}/shares/member`,
      headers: { cookie: adminCookie, 'content-type': 'application/json' },
      payload: { handle: 'carol', role: 'edit' },
    });
    assert.equal(r.statusCode, 201, r.body);
  } finally {
    await cleanup();
  }
});

test('member ACL: anonymous → 401 on every member route', async () => {
  const { app, cleanup } = await makeApp({ mode: 'multi' });
  try {
    for (const route of [
      ['GET', '/files/x/shares/members'],
      ['POST', '/files/x/shares/member'],
      ['PATCH', '/files/x/shares/member/1'],
      ['DELETE', '/files/x/shares/member/1'],
    ] as const) {
      const r = await app.inject({ method: route[0], url: route[1] });
      assert.equal(r.statusCode, 401, `${route[0]} ${route[1]} should 401, got ${r.statusCode}`);
    }
  } finally {
    await cleanup();
  }
});

test('member ACL: single mode 404s the member routes (link-only)', async () => {
  const { app, cleanup } = await makeApp({ mode: 'single' });
  try {
    const aliceCookie = await signup(app, 'alice', 'longpassword');
    const fileId = await uploadFile(app, aliceCookie);
    // The link surface still works in single mode...
    const link = await app.inject({
      method: 'GET',
      url: `/files/${fileId}/shares`,
      headers: { cookie: aliceCookie },
    });
    assert.equal(link.statusCode, 200);
    // ...but member ACLs are multi-only → 404 even for the owner.
    const r = await app.inject({
      method: 'GET',
      url: `/files/${fileId}/shares/members`,
      headers: { cookie: aliceCookie },
    });
    assert.equal(r.statusCode, 404);
  } finally {
    await cleanup();
  }
});

test("member ACL: mode 'none' shadows the member routes with 503", async () => {
  const { app, cleanup } = await makeApp({ mode: 'none' });
  try {
    const r = await app.inject({ method: 'GET', url: '/files/x/shares/members' });
    assert.equal(r.statusCode, 503);
  } finally {
    await cleanup();
  }
});
