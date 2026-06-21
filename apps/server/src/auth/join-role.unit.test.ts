/**
 * Branch-matrix unit test for `resolveJoinRole` — the PURE, security-
 * critical decision behind the collab join handshake (sharing-model
 * §6.1 enforcement). This is the most important test surface in the
 * batch: `yjs.ts` onAuthenticate is a thin adapter over this function,
 * so exercising every branch here is what proves the gate is correct.
 *
 * No DB, no socket — the token lookup + password comparator are
 * injected, so we drive every path deterministically.
 *
 * Run with `pnpm test:unit`.
 */
import { strict as assert } from 'node:assert';
import { test } from 'node:test';
import bcrypt from 'bcryptjs';

import { resolveJoinRole } from './join-role.js';
import type { ShareLinkRole } from './personal.js';

const ROOM = 'room-xyz';

/** Build a ShareLinkRole the injected lookup returns. */
function link(overrides: Partial<ShareLinkRole> = {}): ShareLinkRole {
  return {
    workbookId: 'f-1',
    roomId: ROOM,
    role: 'edit',
    hasPassword: false,
    passwordHash: null,
    expiresAt: null,
    ...overrides,
  };
}

/** A lookup that returns the given link for ANY token (or null). */
const lookupReturns =
  (value: ShareLinkRole | null) =>
  (_token: string): ShareLinkRole | null =>
    value;

test('no token → anonymous fall-through (legacy behaviour preserved)', () => {
  for (const token of [null, '']) {
    const r = resolveJoinRole({
      token,
      documentName: ROOM,
      sharePassword: null,
      lookup: () => {
        throw new Error('lookup must NOT be called when there is no token');
      },
    });
    assert.deepEqual(r, { via: 'anonymous' }, `token=${JSON.stringify(token)}`);
  }
});

test('valid view token → read-only, role view, via share-token', () => {
  const r = resolveJoinRole({
    token: 'tok',
    documentName: ROOM,
    sharePassword: null,
    lookup: lookupReturns(link({ role: 'view' })),
  });
  assert.deepEqual(r, { readOnly: true, role: 'view', via: 'share-token' });
});

test('valid edit token → writable, role edit', () => {
  const r = resolveJoinRole({
    token: 'tok',
    documentName: ROOM,
    sharePassword: null,
    lookup: lookupReturns(link({ role: 'edit' })),
  });
  assert.deepEqual(r, { readOnly: false, role: 'edit', via: 'share-token' });
});

test('comment role collapses to read-only (binary readOnly; fine-grained deferred)', () => {
  const r = resolveJoinRole({
    token: 'tok',
    documentName: ROOM,
    sharePassword: null,
    lookup: lookupReturns(link({ role: 'comment' })),
  });
  assert.deepEqual(r, { readOnly: true, role: 'comment', via: 'share-token' });
});

test('unknown / expired token → reject invalid-token', () => {
  // getLinkRole returns null for both unknown AND expired, so a single
  // null-lookup covers both — the gate can't (and need not) tell them
  // apart.
  const r = resolveJoinRole({
    token: 'tok',
    documentName: ROOM,
    sharePassword: null,
    lookup: lookupReturns(null),
  });
  assert.deepEqual(r, { reject: 'invalid-token' });
});

test('token bound to a different room → reject room-mismatch (replay guard)', () => {
  const r = resolveJoinRole({
    token: 'tok',
    documentName: ROOM,
    sharePassword: null,
    lookup: lookupReturns(link({ roomId: 'some-other-room' })),
  });
  assert.deepEqual(r, { reject: 'room-mismatch' });
});

test('legacy empty-roomId token never matches a real room → reject room-mismatch', () => {
  const r = resolveJoinRole({
    token: 'tok',
    documentName: ROOM,
    sharePassword: null,
    lookup: lookupReturns(link({ roomId: '' })),
  });
  assert.deepEqual(r, { reject: 'room-mismatch' });
});

test('password-gated token, no share password supplied → reject password-required', () => {
  const hash = bcrypt.hashSync('secret', 10);
  for (const sp of [null, '']) {
    const r = resolveJoinRole({
      token: 'tok',
      documentName: ROOM,
      sharePassword: sp,
      lookup: lookupReturns(link({ hasPassword: true, passwordHash: hash })),
    });
    assert.deepEqual(r, { reject: 'password-required' }, `sp=${JSON.stringify(sp)}`);
  }
});

test('password-gated token, wrong share password → reject password-mismatch', () => {
  const hash = bcrypt.hashSync('secret', 10);
  const r = resolveJoinRole({
    token: 'tok',
    documentName: ROOM,
    sharePassword: 'WRONG',
    lookup: lookupReturns(link({ hasPassword: true, passwordHash: hash })),
  });
  assert.deepEqual(r, { reject: 'password-mismatch' });
});

test('password-gated token, correct share password → authorised at the token role', () => {
  const hash = bcrypt.hashSync('secret', 10);
  const r = resolveJoinRole({
    token: 'tok',
    documentName: ROOM,
    sharePassword: 'secret',
    lookup: lookupReturns(link({ role: 'view', hasPassword: true, passwordHash: hash })),
  });
  assert.deepEqual(r, { readOnly: true, role: 'view', via: 'share-token' });
});

test('room check runs BEFORE the password check (wrong room never reveals password validity)', () => {
  // A token bound to another room must reject on room-mismatch even when
  // a correct password is supplied — order matters so we never run the
  // (costly + information-leaking) bcrypt compare for a token that isn't
  // even for this room.
  const hash = bcrypt.hashSync('secret', 10);
  let compareCalled = false;
  const r = resolveJoinRole({
    token: 'tok',
    documentName: ROOM,
    sharePassword: 'secret',
    lookup: lookupReturns(link({ roomId: 'other', hasPassword: true, passwordHash: hash })),
    comparePassword: () => {
      compareCalled = true;
      return true;
    },
  });
  assert.deepEqual(r, { reject: 'room-mismatch' });
  assert.equal(compareCalled, false, 'password comparator must not run for a wrong-room token');
});

test('injected comparator is used (custom hash verification path)', () => {
  const r = resolveJoinRole({
    token: 'tok',
    documentName: ROOM,
    sharePassword: 'anything',
    lookup: lookupReturns(link({ role: 'edit', hasPassword: true, passwordHash: 'opaque' })),
    comparePassword: (plain, hashArg) => plain === 'anything' && hashArg === 'opaque',
  });
  assert.deepEqual(r, { readOnly: false, role: 'edit', via: 'share-token' });
});
