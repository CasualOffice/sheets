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
 * Lock the parser shape for the personal-mode IA. Pure function tests —
 * no DOM, no React, no history API.
 */

import { describe, it } from 'node:test';
import assert from 'node:assert/strict';

import { parseRoute } from './index';

describe('parseRoute', () => {
  it('home: `/home`', () => {
    const r = parseRoute('/home', '');
    assert.equal(r.kind, 'home');
    assert.equal(r.shareToken, null);
  });

  it('home: `/` (canonical redirect handled by useRoute)', () => {
    const r = parseRoute('/', '');
    assert.equal(r.kind, 'home');
  });

  it('templates: `/templates`', () => {
    const r = parseRoute('/templates', '');
    assert.equal(r.kind, 'templates');
  });

  it('sheet with simple id', () => {
    const r = parseRoute('/sheet/f_abc123', '');
    assert.equal(r.kind, 'sheet');
    assert.equal(r.id, 'f_abc123');
    assert.equal(r.shareToken, null);
  });

  it('sheet with URL-encoded id', () => {
    const r = parseRoute('/sheet/' + encodeURIComponent('Q3 P&L'), '');
    assert.equal(r.kind, 'sheet');
    assert.equal(r.id, 'Q3 P&L');
  });

  it('sheet-draft: `/sheet/new`', () => {
    const r = parseRoute('/sheet/new', '');
    assert.equal(r.kind, 'sheet-draft');
    assert.equal(r.id, '');
  });

  it('sheet with share token', () => {
    const r = parseRoute('/sheet/f_xyz', '?share=tok123');
    assert.equal(r.kind, 'sheet');
    assert.equal(r.id, 'f_xyz');
    assert.equal(r.shareToken, 'tok123');
  });

  it('room: legacy `/r/<roomId>` still recognised', () => {
    const r = parseRoute('/r/room42', '');
    assert.equal(r.kind, 'room');
    assert.equal(r.id, 'room42');
  });

  it('unknown: arbitrary path falls through', () => {
    const r = parseRoute('/something/else', '');
    assert.equal(r.kind, 'unknown');
  });

  it('search is preserved on every kind', () => {
    const r = parseRoute('/home', '?disableX=1');
    assert.equal(r.search, '?disableX=1');
  });
});
