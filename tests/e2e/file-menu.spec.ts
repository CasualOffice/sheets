import { expect, test } from '@playwright/test';
import { waitForUniver } from './_helpers';

test.describe('File menu', () => {
  test('Clicking File opens an Office 365-style dropdown', async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);

    await expect(page.getByTestId('file-menu')).toHaveCount(0);
    await page.getByTestId('ribbon-tab-file').click();

    await expect(page.getByTestId('file-menu')).toBeVisible();
    await expect(page.getByTestId('file-menu-properties')).toBeVisible();
  });

  test('Escape closes the File menu', async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);

    await page.getByTestId('ribbon-tab-file').click();
    await expect(page.getByTestId('file-menu')).toBeVisible();

    await page.keyboard.press('Escape');
    await expect(page.getByTestId('file-menu')).toHaveCount(0);
  });
});

test.describe('Properties dialog', () => {
  test('Opens, shows computed fields, and persists edited metadata', async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);

    // Add some content so the cell count is non-zero.
    await page.evaluate(() => {
      const api = window.__univerAPI!;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.getActiveWorkbook()!.getActiveSheet();
      ws.getRange('A1').setValue({ v: 'one' });
      ws.getRange('B2').setValue({ v: 'two' });
      ws.getRange('C3').setValue({ v: 'three' });
    });

    await page.getByTestId('ribbon-tab-file').click();
    await page.getByTestId('file-menu-properties').click();

    await expect(page.getByTestId('properties-dialog')).toBeVisible();
    await expect(page.getByTestId('prop-sheets')).toHaveText('1');
    await expect(page.getByTestId('prop-cells')).toHaveText('3');
    await expect(page.getByTestId('prop-size')).toHaveText(/B|KB/);

    await page.getByTestId('prop-title').fill('Q4 Forecast');
    await page.getByTestId('prop-author').fill('Sachin');
    await page.getByTestId('prop-tags').fill('finance, draft');
    await page.getByTestId('prop-description').fill('Initial forecast for Q4.');

    await page.getByTestId('properties-save').click();
    await expect(page.getByTestId('properties-dialog')).toHaveCount(0);

    // Re-open and verify persistence.
    await page.getByTestId('ribbon-tab-file').click();
    await page.getByTestId('file-menu-properties').click();
    await expect(page.getByTestId('prop-title')).toHaveValue('Q4 Forecast');
    await expect(page.getByTestId('prop-author')).toHaveValue('Sachin');
    await expect(page.getByTestId('prop-tags')).toHaveValue('finance, draft');
    await expect(page.getByTestId('prop-description')).toHaveValue('Initial forecast for Q4.');
  });

  test('Escape closes the dialog', async ({ page }) => {
    await page.goto('/');
    await waitForUniver(page);

    await page.getByTestId('ribbon-tab-file').click();
    await page.getByTestId('file-menu-properties').click();
    await expect(page.getByTestId('properties-dialog')).toBeVisible();

    await page.keyboard.press('Escape');
    await expect(page.getByTestId('properties-dialog')).toHaveCount(0);
  });
});
