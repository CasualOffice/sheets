import { expect, test } from '@playwright/test';
import { waitForUniver } from './_helpers';

/**
 * @mention data source — Phase 3, T3.1/T3.3 foundation.
 *
 * Univer's comment editor lists @-mention candidates via the core
 * `IMentionIOService`. The default is hardwired to the current user; CasualSheets
 * overrides it (at `new Univer`) with the host-pluggable `CasualMentionIOService`
 * (see packages/sdk/src/sheets/mention-source.ts). This verifies, end-to-end
 * through the live Univer injector, that:
 *   - the DI override is in effect (the service resolves to ours, not the default
 *     single-current-user list), and
 *   - it reads the installed provider and applies the search filter.
 * The in-editor @-autocomplete popup rides on top of this exact service.
 */
test('the live mention service resolves to the pluggable provider', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);

  const labels = await page.evaluate(async () => {
    window.__setMentionProvider__?.([
      { id: 'self:Ada', label: 'Ada Lovelace' },
      { id: 'peer:1', label: 'Grace Hopper' },
      { id: 'peer:2', label: 'Bo Diddley' },
    ]);
    return window.__mentionList__?.('');
  });

  // Default service would return only the (empty) current user; ours returns
  // every installed candidate → proof the override took.
  expect(labels).toEqual(['Ada Lovelace', 'Grace Hopper', 'Bo Diddley']);
});

test('the mention service applies the @-search filter', async ({ page }) => {
  await page.goto('/');
  await waitForUniver(page);

  const filtered = await page.evaluate(async () => {
    window.__setMentionProvider__?.([
      { id: 'a', label: 'Ada Lovelace' },
      { id: 'b', label: 'Grace Hopper' },
    ]);
    return window.__mentionList__?.('hop');
  });

  expect(filtered).toEqual(['Grace Hopper']);
});
