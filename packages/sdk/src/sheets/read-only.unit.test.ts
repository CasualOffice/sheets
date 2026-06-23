/**
 * Unit tests for the comment-only veto predicate (the share-role `comment`
 * security boundary). Runs under `node --import tsx` via `node:test`.
 *
 * The stateful veto wiring (beforeCommandExecuted on a real injector) is
 * exercised through the collab flow; this file pins the pure allow/block
 * contract — which commands a commenter may run.
 *
 * Run with: `pnpm test:unit`
 */
import { strict as assert } from 'node:assert';
import { test } from 'node:test';

import { isCommentOnlyBlocked } from './read-only-predicate';

test('comment-only BLOCKS cell / style / structural mutations', () => {
  for (const id of [
    'sheet.command.set-range-values',
    'sheet.mutation.set-range-values',
    'sheet.command.set-style',
    'sheet.command.set-bold',
    'sheet.command.insert-row',
    'sheet.command.remove-col',
    'sheet.command.set-cell-edit-visible',
    'sheet.command.paste',
    'sheet.command.move-range',
    'sheet.command.merge-cells',
  ]) {
    assert.equal(isCommentOnlyBlocked(id), true, `expected blocked: ${id}`);
  }
});

test('comment-only ALLOWS all threaded-comment + comment-editor commands', () => {
  for (const id of [
    'thread-comment.command.add-comment',
    'thread-comment.mutation.add-comment',
    'thread-comment.command.update-comment',
    'thread-comment.command.resolve-comment',
    'thread-comment.command.delete-comment', // would match READONLY_BLOCK's `delete-` — explicitly allowed
    'thread-comment.command.delete-comment-tree',
    'doc.mutation.rich-text-editing', // the comment editor's text edits
  ]) {
    assert.equal(isCommentOnlyBlocked(id), false, `expected allowed: ${id}`);
  }
});

test('comment-only ALLOWS navigation / read commands (fall through)', () => {
  for (const id of [
    'sheet.operation.set-selections',
    'sheet.command.copy',
    'univer.command.undo',
    'univer.command.redo',
    'sheet.operation.set-zoom-ratio',
  ]) {
    assert.equal(isCommentOnlyBlocked(id), false, `expected allowed: ${id}`);
  }
});
