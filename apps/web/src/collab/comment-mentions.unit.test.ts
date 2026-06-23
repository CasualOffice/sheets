import { strict as assert } from 'node:assert';
import { test } from 'node:test';

import { commentMentionsName, extractMentionLabels, type MentionBody } from './comment-mentions.js';

// CustomRangeType.MENTION is 6 in core. Passed in so this test needs no
// @univerjs value import (the tsx runner can't load those).
const MENTION = 6;
const HYPERLINK = 0;

// "@Ada Lovelace, see @Bo" — two mentions + a non-mention range.
const BODY: MentionBody = {
  dataStream: '@Ada Lovelace, see @Bo\r\n',
  customRanges: [
    { rangeType: MENTION, startIndex: 0, endIndex: 12 }, // "@Ada Lovelace"
    { rangeType: MENTION, startIndex: 19, endIndex: 21 }, // "@Bo"
    { rangeType: HYPERLINK, startIndex: 5, endIndex: 9 }, // ignored
  ],
};

test('extractMentionLabels returns mention labels with @ stripped', () => {
  assert.deepEqual(extractMentionLabels(BODY, MENTION), ['Ada Lovelace', 'Bo']);
});

test('extractMentionLabels ignores non-mention custom ranges', () => {
  const onlyLink: MentionBody = {
    dataStream: 'see docs',
    customRanges: [{ rangeType: HYPERLINK, startIndex: 0, endIndex: 2 }],
  };
  assert.deepEqual(extractMentionLabels(onlyLink, MENTION), []);
});

test('extractMentionLabels tolerates a missing/empty body', () => {
  assert.deepEqual(extractMentionLabels(undefined, MENTION), []);
  assert.deepEqual(extractMentionLabels({}, MENTION), []);
});

test('commentMentionsName matches case-insensitively', () => {
  assert.equal(commentMentionsName(BODY, MENTION, 'ada lovelace'), true);
  assert.equal(commentMentionsName(BODY, MENTION, 'Bo'), true);
  assert.equal(commentMentionsName(BODY, MENTION, 'Grace Hopper'), false);
});

test('commentMentionsName returns false for an empty/blank name', () => {
  assert.equal(commentMentionsName(BODY, MENTION, ''), false);
  assert.equal(commentMentionsName(BODY, MENTION, '   '), false);
  assert.equal(commentMentionsName(BODY, MENTION, null), false);
});
