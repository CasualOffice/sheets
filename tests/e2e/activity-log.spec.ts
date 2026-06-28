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
 * Activity log — UX_AUDIT.md §4.1 / Phase 4 #14.
 *
 * Drives the end-to-end shape with synthetic errors fired through the
 * window-event bridge (same channel ToastContext uses). Avoids
 * coupling the spec to any specific real failure surface — the
 * bridge is the contract.
 */
import { test, expect } from '@playwright/test';
import { waitForUniver } from './_helpers';

test.describe('Activity log', () => {
  test('idle: no pill rendered when there are no entries', async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);
    await expect(page.getByTestId('activity-pill')).toHaveCount(0);
  });

  test('error event surfaces a badge; popover lists it; dismiss removes it', async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);

    // Fire two synthetic error events through the same bridge the
    // ToastContext uses on every toast.error call.
    await page.evaluate(() => {
      window.dispatchEvent(
        new CustomEvent('cd:activity-error', {
          detail: { message: 'Save failed: network down' },
        }),
      );
      window.dispatchEvent(
        new CustomEvent('cd:activity-error', {
          detail: { message: 'Open failed: bad xlsx' },
        }),
      );
    });

    // Pill appears with badge "2".
    await expect(page.getByTestId('activity-pill')).toBeVisible();
    const badge = page.getByTestId('activity-pill-badge');
    await expect(badge).toHaveText('2');

    // Open the popover; the badge clears (markAllRead).
    await page.getByTestId('activity-pill-trigger').click();
    await expect(page.getByTestId('activity-pill-popover')).toBeVisible();
    await expect(page.getByText('Save failed: network down')).toBeVisible();
    await expect(page.getByText('Open failed: bad xlsx')).toBeVisible();
    // Badge gone once read.
    await expect(page.getByTestId('activity-pill-badge')).toHaveCount(0);

    // Dismiss one entry — the other survives.
    const firstDismiss = page.locator('[data-testid$="-dismiss"]').first();
    await firstDismiss.click();
    // 2 → 1 entry; the popover header reflects it.
    await expect(page.getByTestId('activity-pill-popover')).toContainText('1 entry');
  });

  test('retryable entry shows a Retry button; click re-runs and success dismisses', async ({
    page,
  }) => {
    await page.goto('/');
    await waitForUniver(page);

    // Push an entry WITH a retry handler via the DEV test seam (the
    // window-event bridge can only carry serializable data, so it can't
    // deliver a closure). The handler fails the first call and succeeds
    // the second, mirroring a transient save failure that recovers.
    await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const w = window as any;
      w.__retryCalls = 0;
      w.__activityRetry__.push(
        "Couldn't save: network down",
        async () => {
          w.__retryCalls += 1;
          if (w.__retryCalls === 1) throw new Error('still down');
          // second call resolves → success
        },
        'save',
      );
    });

    // Pill appears; open it.
    await expect(page.getByTestId('activity-pill')).toBeVisible();
    await page.getByTestId('activity-pill-trigger').click();
    await expect(page.getByTestId('activity-pill-popover')).toBeVisible();
    await expect(page.getByText("Couldn't save: network down")).toBeVisible();

    // Retry button is present for the retryable entry.
    const retry = page.getByTestId('activity-entry-retry');
    await expect(retry).toBeVisible();

    // First click fails — entry survives and a "Retry failed" entry is
    // added, so the original is still there with its Retry button.
    await retry.click();
    await expect(page.getByText('Retry failed: still down')).toBeVisible();
    await expect(page.getByText("Couldn't save: network down")).toBeVisible();

    // Second click succeeds — the original entry is dismissed.
    await page.getByTestId('activity-entry-retry').first().click();
    await expect(page.getByText("Couldn't save: network down")).toHaveCount(0);

    // The handler ran exactly twice.
    const calls = await page.evaluate(() => (window as { __retryCalls?: number }).__retryCalls);
    expect(calls).toBe(2);
  });

  test('Clear all empties the log and removes the pill', async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);

    await page.evaluate(() => {
      window.dispatchEvent(
        new CustomEvent('cd:activity-error', {
          detail: { message: 'Boom' },
        }),
      );
    });

    await page.getByTestId('activity-pill-trigger').click();
    await page.getByTestId('activity-clear-all').click();

    // Pill self-hides once the log is empty.
    await expect(page.getByTestId('activity-pill')).toHaveCount(0);
  });
});
