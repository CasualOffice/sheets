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
 * Desktop local-user chip (slice 5). In the desktop build the collab cluster is
 * hidden; a profile chip stands in, showing who you're signed in as — fetched
 * from the shell's profile.json via `bridge.getProfile()`. Web shows neither.
 */
test('desktop shows a local-user profile chip from the bridge profile', async ({ page }) => {
  test.setTimeout(60_000);
  // Mock the Tauri IPC layer so the real bootstrap builds its bridge and
  // get_profile resolves. document_size=0 → the boot load-effect no-ops.
  await page.addInitScript(() => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (window as any).__TAURI__ = {
      core: {
        invoke: (cmd: string) => {
          if (cmd === 'get_profile')
            return Promise.resolve({ name: 'Grace Hopper', avatar_hue: 200, avatar_path: null });
          if (cmd === 'document_size') return Promise.resolve(0);
          return Promise.resolve(null);
        },
      },
      window: { getCurrentWindow: () => ({ setTitle: () => Promise.resolve() }) },
    };
  });
  await page.goto('/?desk=1');
  await waitForUniver(page);

  const chip = page.getByTestId('titlebar-profile-chip');
  await expect(chip).toBeVisible();
  await expect(chip).toContainText('Grace'); // first name
  await expect(chip.locator('.titlebar__profile-avatar')).toHaveText('GH'); // initials
  // Collab cluster + Share stay hidden in desktop.
  await expect(page.getByTestId('titlebar-collab')).toHaveCount(0);
});

test('web shows no profile chip (control)', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);
  await expect(page.getByTestId('titlebar-profile-chip')).toHaveCount(0);
});
