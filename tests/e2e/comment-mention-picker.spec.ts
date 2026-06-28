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
 * @mention autocomplete in the comment editor — Phase 3, T3.1/T3.3.
 *
 * Typing `@` in the comment composer opens `docs-mention-ui`'s picker, which
 * lists candidates from the core `IMentionIOService` — overridden by CasualSheets
 * with the host-pluggable `CasualMentionIOService`. This drives the real popup:
 * type `@` + a query and assert the matching collaborator appears under the
 * PEOPLE group while the non-match is filtered out. (Selecting a candidate is
 * Univer's own tested behaviour; the picker row sits under the comment editor's
 * pointer-capture overlay, so we don't click-assert insertion here.)
 *
 * Candidates come from the test mention provider (`__setMentionProvider__`); in
 * the app the provider is the live presence peers (CollabDriver). The provider →
 * service path is also covered headless-free in comment-mention-source.spec.ts.
 */
test('typing @ in a comment lists the matching collaborator under PEOPLE', async ({ page }) => {
  test.setTimeout(60_000);
  await page.goto('/');
  await waitForUniver(page);
  await page.evaluate(() => window.__ensurePlugin__?.('threadComment'));

  // Install candidates + open the comment composer on B2.
  await page.evaluate(() => {
    window.__setMentionProvider__?.([
      { id: 'p1', label: 'Grace Hopper' },
      { id: 'p2', label: 'Ada Lovelace' },
    ]);
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange('B2').activate();
    api.executeCommand('sheet.operation.show-comment-modal');
  });

  // The composer auto-focuses its editor. Trigger the mention popup.
  await page.keyboard.type('@');
  await page.keyboard.type('Gr');

  // The PEOPLE group + the matching candidate render; the non-match is filtered.
  await expect(page.getByText('PEOPLE')).toBeVisible({ timeout: 10_000 });
  await expect(page.getByText('Grace Hopper')).toBeVisible();
  await expect(page.getByText('Ada Lovelace')).toHaveCount(0);
});
