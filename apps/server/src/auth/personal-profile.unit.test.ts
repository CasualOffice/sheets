/**
 * Profile + avatar store contract — covers the editable fields
 * (display_name, email, timezone, avatar bytes, preferences JSON)
 * added in Phase C Batch 4.5 of #49.
 */
import { strict as assert } from 'node:assert';
import { test } from 'node:test';
import { mkdtempSync, rmSync } from 'node:fs';
import { join } from 'node:path';
import { tmpdir } from 'node:os';

import { PersonalAuthStore } from './personal';

function withUserStore(): {
  store: PersonalAuthStore;
  userId: number;
  cleanup: () => void;
} {
  const dir = mkdtempSync(join(tmpdir(), 'casual-profile-'));
  const store = new PersonalAuthStore({
    dbPath: join(dir, 'users.db'),
    mode: 'multi',
    bootstrap: null,
  });
  const u = store.createUser('alice', 'longpassword');
  if (!u.ok) throw new Error('seed failed');
  return {
    store,
    userId: u.user.id,
    cleanup: () => {
      store.close();
      rmSync(dir, { recursive: true, force: true });
    },
  };
}

test('getProfile returns sensible defaults for a fresh user', () => {
  const { store, userId, cleanup } = withUserStore();
  try {
    const p = store.getProfile(userId);
    assert.ok(p);
    assert.equal(p?.displayName, null);
    assert.equal(p?.email, null);
    assert.equal(p?.timezone, 'UTC');
    assert.equal(p?.hasAvatar, false);
    assert.deepEqual(p?.preferences, {});
  } finally {
    cleanup();
  }
});

test('updateProfile writes display name + email + timezone + preferences', () => {
  const { store, userId, cleanup } = withUserStore();
  try {
    const updated = store.updateProfile(userId, {
      displayName: 'Alice Wonderland',
      email: 'alice@example.com',
      timezone: 'America/New_York',
      preferences: { theme: 'dark', locale: 'en-US' },
    });
    assert.ok(updated);
    assert.equal(updated?.displayName, 'Alice Wonderland');
    assert.equal(updated?.email, 'alice@example.com');
    assert.equal(updated?.timezone, 'America/New_York');
    assert.deepEqual(updated?.preferences, { theme: 'dark', locale: 'en-US' });
  } finally {
    cleanup();
  }
});

test('updateProfile rejects malformed email', () => {
  const { store, userId, cleanup } = withUserStore();
  try {
    const result = store.updateProfile(userId, { email: 'not-an-email' });
    assert.equal(result, null);
    // Profile stays at defaults.
    assert.equal(store.getProfile(userId)?.email, null);
  } finally {
    cleanup();
  }
});

test('updateProfile clears email when explicitly set to null or empty', () => {
  const { store, userId, cleanup } = withUserStore();
  try {
    store.updateProfile(userId, { email: 'alice@example.com' });
    const cleared = store.updateProfile(userId, { email: '' });
    assert.equal(cleared?.email, null);
  } finally {
    cleanup();
  }
});

test('updateProfile enforces email uniqueness across users', () => {
  const { store, userId, cleanup } = withUserStore();
  try {
    const b = store.createUser('bob', 'longpassword');
    if (!b.ok) throw new Error('seed failed');
    store.updateProfile(userId, { email: 'shared@example.com' });
    const conflict = store.updateProfile(b.user.id, { email: 'shared@example.com' });
    assert.equal(conflict, null);
  } finally {
    cleanup();
  }
});

test('updateProfile is a no-op when patch is empty', () => {
  const { store, userId, cleanup } = withUserStore();
  try {
    const before = store.getProfile(userId);
    const after = store.updateProfile(userId, {});
    assert.deepEqual(after, before);
  } finally {
    cleanup();
  }
});

test('setAvatar persists bytes + mime; getAvatar reads them back', () => {
  const { store, userId, cleanup } = withUserStore();
  try {
    const bytes = new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8]);
    store.setAvatar(userId, 'image/png', bytes);
    const re = store.getAvatar(userId);
    assert.ok(re);
    assert.equal(re?.mime, 'image/png');
    assert.deepEqual(Array.from(re!.bytes), Array.from(bytes));
    assert.equal(store.getProfile(userId)?.hasAvatar, true);
  } finally {
    cleanup();
  }
});

test('setAvatar(null, null) clears the avatar', () => {
  const { store, userId, cleanup } = withUserStore();
  try {
    store.setAvatar(userId, 'image/png', new Uint8Array([1, 2, 3]));
    store.setAvatar(userId, null, null);
    assert.equal(store.getAvatar(userId), null);
    assert.equal(store.getProfile(userId)?.hasAvatar, false);
  } finally {
    cleanup();
  }
});
