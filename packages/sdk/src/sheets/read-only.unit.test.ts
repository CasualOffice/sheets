/**
 * Copyright 2026 Casual Office
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

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

import { isCommentOnlyBlocked, isReadOnlyBlocked } from './read-only-predicate';

test('read-only classification covers real Univer formatting roots', () => {
  for (const id of [
    'sheet.command.set-range-bold',
    'sheet.command.set-range-italic',
    'sheet.command.set-range-font-family',
    'sheet.command.set-range-text-color',
    'sheet.command.reset-range-text-color',
    'sheet.command.numfmt.set.currency',
  ]) {
    assert.equal(isReadOnlyBlocked(id), true, `expected blocked: ${id}`);
  }
});

test('read-only classification fails closed for plugin and future mutators', () => {
  for (const id of [
    'sheet.command.add-table',
    'sheet.command.set-table-config',
    'sheet.command.smart-toggle-filter',
    'sheet.command.set-filter-criteria',
    'sheet.command.addDataValidation',
    'sheet.command.apply-format-painter',
    'sheet.command.refill',
    'sheet.command.copy-down',
    'sheet.command.copy-right',
    'sheet.command.append-row',
    'future-plugin.mutation.rewrite-cells',
  ]) {
    assert.equal(isReadOnlyBlocked(id), true, `expected blocked: ${id}`);
  }
});

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
    'univer.command.undo',
    'univer.command.redo',
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

test('read-only and comment-only allow real navigation, selection, zoom, sheet switch, and copy', () => {
  for (const id of [
    'sheet.operation.set-selections',
    'sheet.command.select-range',
    'sheet.command.move-selection',
    'sheet.command.move-selection-enter-tab',
    'sheet.command.expand-selection',
    'sheet.command.select-all',
    'sheet.command.set-scroll-relative',
    'sheet.command.scroll-view',
    'sheet.command.scroll-to-cell',
    'sheet.command.scroll-view-reset',
    'sheet.command.change-zoom-ratio',
    'sheet.command.set-zoom-ratio',
    'sheet.command.set-worksheet-activate',
    'sheet.command.copy',
    'sheet.command.copy-formula-only',
    'sheet.operation.set-zoom-ratio',
  ]) {
    assert.equal(isReadOnlyBlocked(id), false, `expected read-only allow: ${id}`);
    assert.equal(isCommentOnlyBlocked(id), false, `expected allowed: ${id}`);
  }
});
