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
 * Share dialog comment-role option — Phase 3, T3.4 follow-up. The `comment`
 * link-role is enforced (anonymous `?role=comment` works via applyCommentOnly),
 * but until now there was no UX to *create* a comment link — CreateRoomDialog
 * only offered Edit / View. This verifies the three-way picker renders + is
 * selectable. (The created comment URL itself is `?role=comment`; the full
 * create-room flow is covered by coedit-share.spec.)
 */
test('share dialog offers a Comment role between Edit and View', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);

  await page.getByTestId('menubar-file').click();
  await page.getByTestId('menu-item-start-room').click();
  await expect(page.getByTestId('share-room-dialog')).toBeVisible();

  // All three roles present.
  await expect(page.getByTestId('share-room-role-write')).toBeVisible();
  await expect(page.getByTestId('share-room-role-comment')).toBeVisible();
  await expect(page.getByTestId('share-room-role-view')).toBeVisible();

  // Comment is selectable.
  await page.getByTestId('share-room-role-comment').click();
  await expect(page.getByTestId('share-room-role-comment')).toHaveAttribute(
    'aria-checked',
    'true',
  );
  await expect(page.getByTestId('share-room-role-write')).toHaveAttribute('aria-checked', 'false');
});
