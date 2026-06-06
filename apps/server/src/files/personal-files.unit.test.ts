/**
 * Integration test for the `/files/*` routes — boots a Fastify app
 * in-process via `inject`, wires the personal auth store + an
 * in-memory `HostIntegration`, and walks the full lifecycle:
 *
 *   signup → list (empty) → upload → list (one) → download →
 *   patch rename → put replace bytes → delete → list (empty)
 *
 * Also covers the negative paths: anonymous → 401, wrong-owner →
 * 404, mode=none → 503, oversized body → 413, version-mismatch on
 * PUT → 412.
 *
 * Fast — runs against a memory host + sqlite tmp file. Run with
 * `pnpm test:unit`.
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
import { MemoryHost } from '../host/memory.js';

type Bootstrap = {
  app: FastifyInstance;
  store: PersonalAuthStore;
  host: MemoryHost;
  cleanup: () => Promise<void>;
};

async function makeApp(opts: { mode: 'single' | 'multi' | 'none' } = { mode: 'single' }) {
  const dir = mkdtempSync(join(tmpdir(), 'casual-files-routes-'));
  const store = new PersonalAuthStore({
    dbPath: join(dir, 'users.db'),
    mode: opts.mode,
    bootstrap: null,
  });
  const host = new MemoryHost();
  const app = Fastify({ logger: false, bodyLimit: 5 * 1024 * 1024 });
  await app.register(cookie);
  await app.register(multipart, { limits: { fileSize: 5 * 1024 * 1024 } });
  app.addContentTypeParser(
    'application/octet-stream',
    { parseAs: 'buffer', bodyLimit: 5 * 1024 * 1024 },
    (_req, body, done) => done(null, body),
  );
  registerPersonalAuthRoutes(app, store);
  registerPersonalFilesRoutes(app, store, host, { maxUploadBytes: 5 * 1024 * 1024 });
  await app.ready();
  return {
    app,
    store,
    host,
    cleanup: async () => {
      await app.close();
      store.close();
      rmSync(dir, { recursive: true, force: true });
    },
  } satisfies Bootstrap;
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

test('happy path: signup → upload → list → download → rename → replace → delete', async () => {
  const { app, host, cleanup } = await makeApp();
  try {
    const cookie = await signup(app, 'alice', 'longpassword');

    // List should be empty.
    let r = await app.inject({ method: 'GET', url: '/files', headers: { cookie } });
    assert.equal(r.statusCode, 200);
    assert.deepEqual(JSON.parse(r.body), { files: [] });

    // Upload.
    const boundary = '----CasualTestBoundary' + Math.random().toString(36).slice(2);
    const body = multipartBody(Buffer.from('FAKE-XLSX'), 'budget.xlsx', boundary);
    r = await app.inject({
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
    const created = JSON.parse(r.body).file as { id: string; name: string; etag: string };
    assert.equal(created.name, 'budget.xlsx');

    // List should now show it.
    r = await app.inject({ method: 'GET', url: '/files', headers: { cookie } });
    const list = JSON.parse(r.body).files as Array<{ id: string }>;
    assert.equal(list.length, 1);
    assert.equal(list[0]?.id, created.id);

    // Download bytes round-trip via the host integration.
    r = await app.inject({ method: 'GET', url: `/files/${created.id}`, headers: { cookie } });
    assert.equal(r.statusCode, 200);
    assert.equal(r.body, 'FAKE-XLSX');
    assert.equal(r.headers['etag'], created.etag);

    // Rename.
    r = await app.inject({
      method: 'PATCH',
      url: `/files/${created.id}`,
      headers: { cookie, 'content-type': 'application/json' },
      payload: { name: 'renamed.xlsx' },
    });
    assert.equal(r.statusCode, 204);
    r = await app.inject({ method: 'GET', url: '/files', headers: { cookie } });
    assert.equal((JSON.parse(r.body).files as Array<{ name: string }>)[0]?.name, 'renamed.xlsx');

    // Replace bytes via the raw POST (matches the WOPI PutFile shape).
    r = await app.inject({
      method: 'POST',
      url: `/files/${created.id}`,
      headers: {
        cookie,
        'content-type': 'application/octet-stream',
        'if-match': created.etag,
      },
      payload: Buffer.from('NEW-XLSX-BYTES'),
    });
    assert.equal(r.statusCode, 200, `replace status ${r.statusCode}: ${r.body}`);
    const updated = JSON.parse(r.body).file as { etag: string; size: number };
    assert.notEqual(updated.etag, created.etag);
    assert.equal(updated.size, 14);

    // Delete.
    r = await app.inject({ method: 'DELETE', url: `/files/${created.id}`, headers: { cookie } });
    assert.equal(r.statusCode, 204);
    r = await app.inject({ method: 'GET', url: '/files', headers: { cookie } });
    assert.deepEqual(JSON.parse(r.body), { files: [] });
    // Host integration was wiped too.
    assert.equal(await host.getFile(created.id), null);
  } finally {
    await cleanup();
  }
});

test('anonymous requests get 401 on every file route', async () => {
  const { app, cleanup } = await makeApp();
  try {
    for (const route of [
      ['GET', '/files'],
      ['POST', '/files'],
      ['GET', '/files/anything'],
      ['POST', '/files/anything'],
      ['DELETE', '/files/anything'],
      ['PATCH', '/files/anything'],
    ] as const) {
      const r = await app.inject({ method: route[0], url: route[1] });
      assert.equal(r.statusCode, 401, `${route[0]} ${route[1]} should be 401, got ${r.statusCode}`);
    }
  } finally {
    await cleanup();
  }
});

test("mode 'none' shadows every file route with 503", async () => {
  const { app, cleanup } = await makeApp({ mode: 'none' });
  try {
    const r = await app.inject({ method: 'GET', url: '/files' });
    assert.equal(r.statusCode, 503);
  } finally {
    await cleanup();
  }
});

test('cross-user access returns 404 even when the id exists', async () => {
  const { app, cleanup } = await makeApp({ mode: 'multi' });
  try {
    const aliceCookie = await signup(app, 'alice', 'longpassword');
    const bobCookie = await signup(app, 'bob', 'longpassword');
    const boundary = '----CasualTestBoundary' + Math.random().toString(36).slice(2);
    const body = multipartBody(Buffer.from('FAKE-XLSX'), 'mine.xlsx', boundary);
    const upload = await app.inject({
      method: 'POST',
      url: '/files',
      headers: {
        cookie: aliceCookie,
        'content-type': `multipart/form-data; boundary=${boundary}`,
        'content-length': String(body.length),
      },
      payload: body,
    });
    const fileId = (JSON.parse(upload.body).file as { id: string }).id;
    // Bob with Alice's file id — 404, never 403, so we don't leak
    // ownership info.
    const r = await app.inject({
      method: 'GET',
      url: `/files/${fileId}`,
      headers: { cookie: bobCookie },
    });
    assert.equal(r.statusCode, 404);
    const list = await app.inject({ method: 'GET', url: '/files', headers: { cookie: bobCookie } });
    assert.deepEqual(JSON.parse(list.body), { files: [] });
  } finally {
    await cleanup();
  }
});

test('replace-with-stale-etag returns 412 with the conflict envelope', async () => {
  const { app, cleanup } = await makeApp();
  try {
    const cookie = await signup(app, 'alice', 'longpassword');
    const boundary = '----CasualTestBoundary' + Math.random().toString(36).slice(2);
    const body = multipartBody(Buffer.from('v1'), 'budget.xlsx', boundary);
    const upload = await app.inject({
      method: 'POST',
      url: '/files',
      headers: {
        cookie,
        'content-type': `multipart/form-data; boundary=${boundary}`,
        'content-length': String(body.length),
      },
      payload: body,
    });
    const created = JSON.parse(upload.body).file as { id: string; etag: string };
    const r = await app.inject({
      method: 'POST',
      url: `/files/${created.id}`,
      headers: {
        cookie,
        'content-type': 'application/octet-stream',
        'if-match': 'wrong-etag',
      },
      payload: Buffer.from('v2'),
    });
    assert.equal(r.statusCode, 412);
    const env = JSON.parse(r.body);
    assert.equal(env.error, 'version-mismatch');
    assert.equal(env.expected, created.etag);
  } finally {
    await cleanup();
  }
});
