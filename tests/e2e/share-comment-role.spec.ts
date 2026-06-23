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
