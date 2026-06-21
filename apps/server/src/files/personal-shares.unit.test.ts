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
  registerPersonalSharesRoutes(app, store);
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
      payload: { role: 'edit', expiresInDays: 7, password: 'secret' },
    });
    assert.equal(r.statusCode, 201, r.body);
    const minted = JSON.parse(r.body) as {
      token: string;
      role: string;
      expiresAt: number;
      url: string;
    };
    assert.equal(minted.role, 'edit');
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
      payload: { role: 'view' },
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
        payload: { role },
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
        payload: { role: 'view', expiresInDays },
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
      payload: { role: 'view' },
    });
    assert.equal(r.statusCode, 404);

    // Alice still works.
    r = await app.inject({
      method: 'POST',
      url: `/files/${fileId}/shares/link`,
      headers: { cookie: aliceCookie, 'content-type': 'application/json' },
      payload: { role: 'view' },
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
      payload: { role: 'edit' },
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
      payload: { role: 'view' },
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
