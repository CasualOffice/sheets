import { strict as assert } from 'node:assert';
import { afterEach, test } from 'node:test';

import {
  filterMentionCandidates,
  getMentionProvider,
  resolveMentionList,
  setMentionProvider,
  type MentionCandidate,
} from './mention-source.js';

// MentionType.PERSON is 0 in core. The pure resolver takes it as a param so
// this test needs no @univerjs value import (which the tsx runner can't load).
const PERSON = 0;

const PEERS: MentionCandidate[] = [
  { id: 'u1', label: 'Ada Lovelace' },
  { id: 'u2', label: 'Bo Diddley', icon: 'http://x/a.png' },
  { id: 'u3', label: 'Grace Hopper' },
];

afterEach(() => {
  setMentionProvider(null);
});

test('filterMentionCandidates: empty search returns all', () => {
  assert.deepEqual(filterMentionCandidates(PEERS, ''), PEERS);
  assert.deepEqual(filterMentionCandidates(PEERS, '   '), PEERS);
});

test('filterMentionCandidates: case-insensitive substring match', () => {
  assert.deepEqual(
    filterMentionCandidates(PEERS, 'ace').map((c) => c.id),
    ['u1', 'u3'], // Ada Lovelace, Grace Hopper
  );
  assert.deepEqual(
    filterMentionCandidates(PEERS, 'BO').map((c) => c.id),
    ['u2'],
  );
});

test('filterMentionCandidates: strips the leading @ trigger', () => {
  // docs-mention-ui passes the slice from the @ anchor, e.g. "@gr".
  assert.deepEqual(
    filterMentionCandidates(PEERS, '@gr').map((c) => c.id),
    ['u3'],
  );
  assert.deepEqual(filterMentionCandidates(PEERS, '@'), PEERS);
});

test('resolveMentionList with no provider lists nobody (safe default)', async () => {
  const res = await resolveMentionList({ search: '' }, PERSON);
  assert.equal(res.total, 0);
  assert.equal(res.list[0].mentions.length, 0);
  assert.equal(res.list[0].title, 'PEOPLE');
  assert.equal(res.list[0].type, PERSON);
});

test('resolveMentionList maps provider candidates to PERSON mentions', async () => {
  setMentionProvider(() => PEERS);
  const res = await resolveMentionList({ search: '' }, PERSON);
  assert.equal(res.total, 3);
  const m = res.list[0].mentions;
  assert.equal(m[0].objectId, 'u1');
  assert.equal(m[0].label, 'Ada Lovelace');
  assert.equal(m[0].objectType, PERSON);
  assert.deepEqual(m[1].metadata, { icon: 'http://x/a.png' });
  assert.equal(m[0].metadata, undefined); // no icon → no metadata
});

test('resolveMentionList awaits an async provider', async () => {
  setMentionProvider(async (search) => filterMentionCandidates(PEERS, search));
  const res = await resolveMentionList({ search: 'hopper' }, PERSON);
  assert.equal(res.total, 1);
  assert.equal(res.list[0].mentions[0].objectId, 'u3');
});

test('a throwing provider degrades to an empty list (never breaks the editor)', async () => {
  setMentionProvider(() => {
    throw new Error('directory down');
  });
  const res = await resolveMentionList({ search: 'x' }, PERSON);
  assert.equal(res.total, 0);
});

test('setMentionProvider(null) clears the provider', () => {
  const fn = () => PEERS;
  setMentionProvider(fn);
  assert.equal(getMentionProvider(), fn);
  setMentionProvider(null);
  assert.equal(getMentionProvider(), null);
});
