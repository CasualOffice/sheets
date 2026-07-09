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
 * Unit tests for the MenuBar feature-gating contract — specifically the
 * embed-mode gates that let a host hide the editor's branding surfaces:
 *
 *   features={{ help: false }}     → the whole Help menu is dropped.
 *   features={{ branding: false }} → the "View on GitHub" + "About" links are
 *                                    dropped (from both Help and File), so an
 *                                    embedded sheet shows no editor branding,
 *                                    while Keyboard shortcuts stays in Help.
 *
 * Exercises the pure `computeVisibleMenus` resolver (no DOM needed). The
 * `MENUS` array itself lives in `MenuBar.tsx`, which statically imports
 * `@univerjs/core` values the vendored ESM can't expose as named exports to
 * node's loader — so it can't be imported here. Instead this pins the resolver
 * against a fixture that mirrors the real Help/File branding subtree (kept in
 * sync with MenuBar.tsx: Help menu `feature: 'help'`; About + View-on-GitHub +
 * their separator `feature: 'branding'`; File > About `feature: 'branding'`).
 *
 * Run with: `pnpm test:unit`
 */
import { strict as assert } from 'node:assert';
import { test } from 'node:test';

import { computeVisibleMenus, type MenuDef } from './menu-model';

// A faithful copy of the branding-relevant subtree of `MENUS` in MenuBar.tsx.
const MENUS_FIXTURE: MenuDef[] = [
  {
    id: 'file',
    label: 'File',
    feature: 'file',
    items: [
      { kind: 'item', id: 'properties', label: 'Properties…', dialog: 'properties' },
      {
        kind: 'item',
        id: 'about',
        label: 'About casual sheets',
        dialog: 'about',
        feature: 'branding',
      },
    ],
  },
  {
    id: 'help',
    label: 'Help',
    feature: 'help',
    items: [
      {
        kind: 'item',
        id: 'keyboard-shortcuts',
        label: 'Keyboard shortcuts',
        dialog: 'keyboard-shortcuts',
      },
      { kind: 'separator', id: 'sep-help', feature: 'branding' },
      {
        kind: 'item',
        id: 'about',
        label: 'About casual sheets',
        dialog: 'about',
        feature: 'branding',
      },
      {
        kind: 'item',
        id: 'github',
        label: 'View on GitHub',
        run: () => {},
        feature: 'branding',
      },
    ],
  },
];

// All SDK-built-in dialogs openable — matches standalone chrome (the SDK ships
// the About / Keyboard-shortcuts modals), so branding items are not
// latent-dropped and the feature gate is the only thing that can hide them.
const canOpenAll = () => true;

function findMenu(menus: MenuDef[], id: string) {
  return menus.find((m) => m.id === id);
}

function itemIds(items: { id: string }[]): string[] {
  return items.map((i) => i.id);
}

test('default (features unset): Help menu, About and View-on-GitHub all render', () => {
  const menus = computeVisibleMenus(MENUS_FIXTURE, {}, canOpenAll);
  const help = findMenu(menus, 'help');
  assert.ok(help, 'Help menu should render by default');
  const ids = itemIds(help.items);
  assert.ok(ids.includes('github'), 'View on GitHub should render by default');
  assert.ok(ids.includes('about'), 'Help > About should render by default');
  assert.ok(ids.includes('keyboard-shortcuts'), 'Keyboard shortcuts should render');

  const file = findMenu(menus, 'file');
  assert.ok(file, 'File menu should render by default');
  assert.ok(itemIds(file.items).includes('about'), 'File > About should render by default');
});

test('features={{ help: false }}: the whole Help menu is dropped', () => {
  const menus = computeVisibleMenus(MENUS_FIXTURE, { help: false }, canOpenAll);
  assert.equal(findMenu(menus, 'help'), undefined, 'Help menu must not render when help:false');
  // File menu is unaffected by the help gate.
  assert.ok(findMenu(menus, 'file'), 'File menu still renders when only help is disabled');
});

test('features={{ branding: false }}: GitHub + About vanish, Help stays for shortcuts', () => {
  const menus = computeVisibleMenus(MENUS_FIXTURE, { branding: false }, canOpenAll);

  const help = findMenu(menus, 'help');
  assert.ok(help, 'Help menu still renders (keyboard shortcuts is not branding)');
  const helpIds = itemIds(help.items);
  assert.ok(!helpIds.includes('github'), 'View on GitHub must be hidden when branding:false');
  assert.ok(!helpIds.includes('about'), 'Help > About must be hidden when branding:false');
  assert.ok(helpIds.includes('keyboard-shortcuts'), 'Keyboard shortcuts stays');
  // Trailing separator before the (now removed) branding items must be collapsed.
  assert.ok(!helpIds.includes('sep-help'), 'dangling separator collapses once branding items go');

  const file = findMenu(menus, 'file');
  assert.ok(file, 'File menu still renders');
  assert.ok(
    !itemIds(file.items).includes('about'),
    'File > About must be hidden when branding:false',
  );
});

test('features={{ help: false, branding: false }}: no editor-branded surface at all', () => {
  const menus = computeVisibleMenus(MENUS_FIXTURE, { help: false, branding: false }, canOpenAll);
  assert.equal(findMenu(menus, 'help'), undefined, 'Help menu gone');
  const file = findMenu(menus, 'file');
  assert.ok(file && !itemIds(file.items).includes('about'), 'File > About gone');
  // No visible menu anywhere links to GitHub.
  for (const m of menus) {
    assert.ok(!itemIds(m.items).includes('github'), `menu ${m.id} has no GitHub link`);
  }
});
