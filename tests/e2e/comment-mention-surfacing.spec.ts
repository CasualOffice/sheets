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
 * "Mentions you" surfacing — Phase 3, T3.3.
 *
 * A comment whose body @-references the local user gets a "You" badge + accent
 * in the comments panel, so you can spot threads that need your attention. The
 * match is by display name (the shared cross-peer identity); mentions are read
 * from the comment body's MENTION custom ranges (comment-mentions.ts).
 */
test('a comment that @mentions me is badged in the panel; others are not', async ({ page }) => {
  test.setTimeout(60_000);
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
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    // A1: @-mentions Ada Lovelace. "@Ada Lovelace" spans indices 0..12.
    await ws.getRange('A1').addComment({
      dataStream: '@Ada Lovelace please review\r\n',
      customRanges: [
        {
          rangeId: 'm1',
          rangeType: 6 /* CustomRangeType.MENTION */,
          startIndex: 0,
          endIndex: 12,
          properties: {},
          wholeEntity: true,
        },
      ],
    });
    // B2: a plain comment, nobody mentioned.
    await ws.getRange('B2').addComment({ dataStream: 'just a note\r\n' });
  });

  await page.getByTestId('panel-rail-comments').click();
  const panel = page.getByTestId('comments-panel');
  await expect(panel).toBeVisible();

  // Exactly one row badged, and it's the one that mentions Ada.
  await expect(panel.locator('.comments-panel__mention-badge')).toHaveCount(1);
  const mentionRow = panel.locator('.comments-panel__row--mentions-me');
  await expect(mentionRow).toHaveCount(1);
  await expect(mentionRow).toContainText('please review');
});
