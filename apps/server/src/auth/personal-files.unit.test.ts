/**
 * Files-registry contract — covers the CRUD on the `files` table in
 * the personal auth store. The route layer composes this with the
 * HostIntegration backend; both are independently tested.
 */
import { strict as assert } from 'node:assert';
import { test } from 'node:test';
import { mkdtempSync, rmSync } from 'node:fs';
import { join } from 'node:path';
import { tmpdir } from 'node:os';

import { PersonalAuthStore } from './personal';

function withStoreUser(): {
  store: PersonalAuthStore;
  userId: number;
  cleanup: () => void;
} {
  const dir = mkdtempSync(join(tmpdir(), 'casual-files-'));
  const store = new PersonalAuthStore({
    dbPath: join(dir, 'users.db'),
    mode: 'multi',
    bootstrap: null,
  });
  const seed = store.createUser('alice', 'longpassword');
  if (!seed.ok) throw new Error('seed failed');
  return {
    store,
    userId: seed.user.id,
    cleanup: () => {
      store.close();
      rmSync(dir, { recursive: true, force: true });
    },
  };
}

test('createFile + getFile round-trip', () => {
  const { store, userId, cleanup } = withStoreUser();
  try {
    const rec = store.createFile({
      ownerId: userId,
      displayName: 'budget.xlsx',
      size: 1024,
      etag: 'v1',
    });
    assert.ok(rec.id.startsWith('f-'));
    const re = store.getFile(rec.id);
    assert.ok(re);
    assert.equal(re?.ownerId, userId);
    assert.equal(re?.displayName, 'budget.xlsx');
    assert.equal(re?.size, 1024);
    assert.equal(re?.etag, 'v1');
  } finally {
    cleanup();
  }
});

test('listFilesForUser returns only the owner’s files, newest-edit first', async () => {
  const { store, userId, cleanup } = withStoreUser();
  try {
    const other = store.createUser('bob', 'longpassword');
    if (!other.ok) throw new Error('seed failed');
    const otherId = other.user.id;
    const a = store.createFile({ ownerId: userId, displayName: 'a.xlsx', size: 1, etag: 'v1' });
    // Small sleep so modified_at is strictly later for `b`.
    await new Promise((r) => setTimeout(r, 5));
    const b = store.createFile({ ownerId: userId, displayName: 'b.xlsx', size: 2, etag: 'v1' });
    store.createFile({ ownerId: otherId, displayName: 'c.xlsx', size: 3, etag: 'v1' });
    const list = store.listFilesForUser(userId);
    assert.equal(list.length, 2);
    assert.equal(list[0]?.id, b.id);
    assert.equal(list[1]?.id, a.id);
  } finally {
    cleanup();
  }
});

test('recordFileUpdate bumps size + etag + modified-at', async () => {
  const { store, userId, cleanup } = withStoreUser();
  try {
    const rec = store.createFile({
      ownerId: userId,
      displayName: 'x.xlsx',
      size: 10,
      etag: 'v1',
    });
    const before = store.getFile(rec.id);
    await new Promise((r) => setTimeout(r, 5));
    store.recordFileUpdate(rec.id, 99, 'v2');
    const after = store.getFile(rec.id);
    assert.equal(after?.size, 99);
    assert.equal(after?.etag, 'v2');
    assert.ok(after && before && after.modifiedAt > before.modifiedAt);
    // created_at stays put
    assert.equal(after?.createdAt, before?.createdAt);
  } finally {
    cleanup();
  }
});

test('renameFile updates display name; rejects empty', () => {
  const { store, userId, cleanup } = withStoreUser();
  try {
    const rec = store.createFile({
      ownerId: userId,
      displayName: 'old.xlsx',
      size: 1,
      etag: 'v1',
    });
    assert.equal(store.renameFile(rec.id, 'new.xlsx'), true);
    assert.equal(store.getFile(rec.id)?.displayName, 'new.xlsx');
    assert.equal(store.renameFile(rec.id, '   '), false);
    assert.equal(store.renameFile('unknown-id', 'anything'), false);
  } finally {
    cleanup();
  }
});

test('deleteFile removes the row + reports the change', () => {
  const { store, userId, cleanup } = withStoreUser();
  try {
    const rec = store.createFile({
      ownerId: userId,
      displayName: 'gone.xlsx',
      size: 1,
      etag: 'v1',
    });
    assert.equal(store.deleteFile(rec.id), true);
    assert.equal(store.getFile(rec.id), null);
    assert.equal(store.deleteFile(rec.id), false);
  } finally {
    cleanup();
  }
});

test('deleting a user cascades to their files', () => {
  const { store, cleanup } = withStoreUser();
  try {
    // `alice` (created by withStoreUser) is the admin and the last
    // admin guard refuses to drop her. Create + delete a regular
    // user instead — same cascade path, exercised through the
    // path that the API actually allows.
    const second = store.createUser('regular', 'longpassword');
    if (!second.ok) throw new Error('seed failed');
    const regularId = second.user.id;
    store.createFile({ ownerId: regularId, displayName: 'mine.xlsx', size: 1, etag: 'v1' });
    assert.equal(store.deleteUser(regularId), true);
    assert.equal(store.listFilesForUser(regularId).length, 0);
  } finally {
    cleanup();
  }
});

test('newFileId issues distinct opaque ids', () => {
  const seen = new Set<string>();
  for (let i = 0; i < 100; i++) {
    seen.add(PersonalAuthStore.newFileId());
  }
  assert.equal(seen.size, 100);
});
