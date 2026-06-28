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

import { expect, test } from '@playwright/test';
import { waitForUniver } from './_helpers';

/**
 * Comment authorship byline — Phase 3, T3.1.
 *
 * Univer stamps a comment's `personId` from the current user, but we can't set
 * a distinct current user per client (setCurrentUser flips the collab grid to
 * read-only — the #122 regression). So authorship is recorded out-of-band: the
 * add-comment *command* runs only on the author's client, where CollabDriver's
 * stamping hook records `commentId → {name,color}` from the local presence
 * identity (and, in a room, mirrors it over a Y.Map for peers).
 *
 * This covers the single-player path: a stored display name becomes the byline
 * on a freshly-created comment. The cross-peer Y.Map sync rides the coedit
 * specs' infrastructure and is exercised there.
 */
test('a new comment shows the local user as its author', async ({ page }) => {
  test.setTimeout(60_000);

  // Seed the display name before the app boots — CollabDriver reads it via
  // getDisplayName() to stamp authorship even outside a room.
  await page.addInitScript(() => {
    localStorage.setItem('casual.collab.displayName', 'Ada Lovelace');
    localStorage.setItem('casual.collab.namePrompted', '1');
  });

  await page.goto('/');
  await waitForUniver(page);
  await page.evaluate(() => window.__ensurePlugin__?.('threadComment'));

  await page.evaluate(async () => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws = api.getActiveWorkbook()!.getActiveSheet() as any;
    await ws.getRange('A1').addComment({ dataStream: 'first pass looks good\r\n' });
  });

  // Open the comments task pane from the panel rail.
  await page.getByTestId('panel-rail-comments').click();
  const panel = page.getByTestId('comments-panel');
  await expect(panel).toBeVisible();

  // The byline resolves the stored identity: name + colored initial avatar.
  await expect(panel.getByText('Ada Lovelace')).toBeVisible();
  await expect(panel.locator('.comments-panel__author-avatar')).toHaveText('AL');
});
