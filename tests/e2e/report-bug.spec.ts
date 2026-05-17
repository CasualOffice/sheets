import { expect, test } from '@playwright/test';
import { waitForUniver } from './_helpers';

/**
 * Help → Report a bug opens the GitHub issue form prefilled with the user's
 * environment so they don't have to type out browser version etc.
 *
 * The structured template lives at .github/ISSUE_TEMPLATE/bug.yml; GitHub
 * resolves `?template=bug.yml&env=…&url=…` to the matching form fields by
 * the `id` declared in the YAML.
 */
test('Help → Report a bug opens GitHub issue form with env prefilled', async ({ page, context }) => {
  await page.goto('/');
  await waitForUniver(page);

  await page.getByTestId('menubar-help').click();
  const popupPromise = context.waitForEvent('page');
  await page.getByTestId('menu-item-report-bug').click();
  const popup = await popupPromise;

  // GitHub may redirect anonymous users to the login page; the original
  // intent survives in `return_to=…`. Decode either layer and assert the
  // template + prefilled fields are all in there.
  const raw = popup.url();
  const inspected = (() => {
    const redirected = new URL(raw);
    const returnTo = redirected.searchParams.get('return_to');
    return returnTo ? decodeURIComponent(returnTo) : raw;
  })();

  expect(inspected).toContain('github.com/schnsrw/sheets/issues/new');
  expect(inspected).toContain('template=bug.yml');
  expect(inspected).toContain('env=');
  expect(inspected).toContain('url=');
  await popup.close();
});
