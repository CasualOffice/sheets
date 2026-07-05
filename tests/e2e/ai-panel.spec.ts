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
 * AI panel visual + behavioural contract.
 *
 * No live LLM calls — tests cover:
 *   - PanelRail toggle (open / close)
 *   - API-key setup view (DirectTransport, no key stored)
 *   - Chat view (key pre-seeded via localStorage)
 *   - Panel mutex (opening another panel closes AI)
 *   - Keyboard close (Escape handled by the close button)
 */

import { expect, test } from '@playwright/test';
import { waitForUniver } from './_helpers';

const AI_KEY = 'casual_sheets_ai_key';
const FAKE_KEY = 'sk-ant-test-key-for-visual-verify';

test.describe('AI panel', () => {
  test.beforeEach(async ({ page }) => {
    // Opt the AI panel into the plain-web (DirectTransport) flow so the rail
    // button is shown — production hides it unless desktop / collab / this flag.
    await page.addInitScript(() => {
      (window as unknown as { __ENABLE_AI__: string }).__ENABLE_AI__ = 'true';
    });
    await page.goto('/');
    await waitForUniver(page);
  });

  // ── Toggle ────────────────────────────────────────────────────────────────

  test('panel is hidden by default', async ({ page }) => {
    await expect(page.getByTestId('ai-panel')).toHaveCount(0);
    await expect(page.getByTestId('panel-rail-ai')).toBeVisible();
    await expect(page.getByTestId('panel-rail-ai')).toHaveAttribute('aria-pressed', 'false');
  });

  test('opens and closes via PanelRail button', async ({ page }) => {
    const btn = page.getByTestId('panel-rail-ai');

    await btn.click();
    await expect(page.getByTestId('ai-panel')).toBeVisible();
    await expect(btn).toHaveAttribute('aria-pressed', 'true');

    await page.screenshot({ path: 'screenshots/ai-panel-open-key-setup.png', fullPage: false });

    await btn.click();
    await expect(page.getByTestId('ai-panel')).toHaveCount(0);
    await expect(btn).toHaveAttribute('aria-pressed', 'false');
  });

  test('close button dismisses the panel', async ({ page }) => {
    await page.getByTestId('panel-rail-ai').click();
    await expect(page.getByTestId('ai-panel')).toBeVisible();

    await page.getByRole('button', { name: /close ai panel/i }).click();
    await expect(page.getByTestId('ai-panel')).toHaveCount(0);
  });

  // ── API-key setup view (DirectTransport, no key in localStorage) ──────────

  test('shows key-setup view when no API key is stored', async ({ page }) => {
    await page.getByTestId('panel-rail-ai').click();
    const panel = page.getByTestId('ai-panel');
    await expect(panel).toBeVisible();

    // Header
    await expect(panel.locator('.side-panel__title')).toContainText('AI');

    // Key-setup copy
    await expect(panel).toContainText(/enter your anthropic api key/i);
    await expect(panel.getByPlaceholder(/sk-ant/i)).toBeVisible();
    await expect(panel.getByRole('button', { name: /save key/i })).toBeVisible();

    await page.screenshot({ path: 'screenshots/ai-panel-key-setup.png', fullPage: false });
  });

  test('saves API key and transitions to chat view', async ({ page }) => {
    await page.getByTestId('panel-rail-ai').click();
    const panel = page.getByTestId('ai-panel');

    await panel.getByPlaceholder(/sk-ant/i).fill(FAKE_KEY);
    await panel.getByRole('button', { name: /save key/i }).click();

    // Key-setup view gone; chat view visible
    await expect(panel).not.toContainText(/enter your anthropic api key/i);
    await expect(panel.getByPlaceholder(/ask about this spreadsheet/i)).toBeVisible();

    // "Change API key" link visible (DirectTransport)
    await expect(panel.getByRole('button', { name: /change api key/i })).toBeVisible();

    await page.screenshot({ path: 'screenshots/ai-panel-chat-empty.png', fullPage: false });
  });

  // ── Chat view with key pre-seeded ─────────────────────────────────────────

  test('chat view — empty state placeholder and input', async ({ page }) => {
    await page.evaluate(([k, v]) => localStorage.setItem(k, v), [AI_KEY, FAKE_KEY] as [
      string,
      string,
    ]);
    await page.reload();
    await waitForUniver(page);

    await page.getByTestId('panel-rail-ai').click();
    const panel = page.getByTestId('ai-panel');
    await expect(panel).toBeVisible();

    // Empty-state prompt
    await expect(panel).toContainText(/ask anything about this spreadsheet/i);

    // Textarea present and enabled
    const textarea = panel.getByRole('textbox', { name: /message input/i });
    await expect(textarea).toBeVisible();
    await expect(textarea).toBeEnabled();

    // Send button present
    await expect(panel.getByRole('button', { name: /send/i })).toBeVisible();

    await page.screenshot({ path: 'screenshots/ai-panel-chat-view.png', fullPage: false });
  });

  test('Enter key in textarea does not add a newline', async ({ page }) => {
    await page.evaluate(([k, v]) => localStorage.setItem(k, v), [AI_KEY, FAKE_KEY] as [
      string,
      string,
    ]);
    await page.reload();
    await waitForUniver(page);

    await page.getByTestId('panel-rail-ai').click();
    const panel = page.getByTestId('ai-panel');
    const textarea = panel.getByRole('textbox', { name: /message input/i });

    await textarea.fill('Hello');
    // Enter without Shift should attempt send (not add newline)
    // The network call will fail (fake key) — we just verify the input clears.
    // Intercept the fetch so it doesn't actually reach Anthropic.
    await page.route('**/v1/messages', (route) =>
      route.fulfill({
        status: 401,
        body: JSON.stringify({ error: { message: 'invalid_api_key' } }),
      }),
    );
    await textarea.press('Enter');
    // The textarea should be cleared (value sent to chat flow)
    await expect(textarea).toHaveValue('');
  });

  // ── Panel mutex ───────────────────────────────────────────────────────────

  test('opening AI panel closes other open panels', async ({ page }) => {
    // Open the Tables panel first
    await page.getByTestId('panel-rail-tables').click();
    await expect(page.getByTestId('tables-panel')).toBeVisible();

    // Opening AI should close Tables
    await page.getByTestId('panel-rail-ai').click();
    await expect(page.getByTestId('ai-panel')).toBeVisible();
    await expect(page.getByTestId('tables-panel')).toHaveCount(0);

    await page.screenshot({ path: 'screenshots/ai-panel-mutex.png', fullPage: false });
  });

  test('opening another panel closes AI panel', async ({ page }) => {
    await page.getByTestId('panel-rail-ai').click();
    await expect(page.getByTestId('ai-panel')).toBeVisible();

    // Opening Tables should close AI
    await page.getByTestId('panel-rail-tables').click();
    await expect(page.getByTestId('tables-panel')).toBeVisible();
    await expect(page.getByTestId('ai-panel')).toHaveCount(0);
  });

  // ── Agentic + MCP ───────────────────────────────────────────────────────────

  test('agent-mode toggle renders and flips Chat ↔ Agent', async ({ page }) => {
    await page.evaluate(([k, v]) => localStorage.setItem(k, v), [AI_KEY, FAKE_KEY] as [
      string,
      string,
    ]);
    await page.reload();
    await waitForUniver(page);
    await page.getByTestId('panel-rail-ai').click();
    await expect(page.getByTestId('ai-panel')).toBeVisible();

    const toggle = page.getByTestId('ai-agent-toggle');
    await expect(toggle).toBeVisible();
    await expect(toggle).toHaveText(/Chat/);
    await toggle.click();
    await expect(toggle).toHaveText(/Agent/);
  });

  test('connects an external MCP server and shows its tool count', async ({ page }) => {
    // Mock a Streamable-HTTP MCP server: reply to initialize + tools/list.
    await page.route('**/mcp-test/rpc', async (route) => {
      const body = JSON.parse(route.request().postData() || '{}');
      if (body.id === undefined) {
        await route.fulfill({ status: 202, body: '' });
        return;
      }
      const result =
        body.method === 'initialize'
          ? {
              protocolVersion: '2025-06-18',
              capabilities: { tools: {} },
              serverInfo: { name: 'mock', version: '1' },
            }
          : body.method === 'tools/list'
            ? {
                tools: [
                  {
                    name: 'web_search',
                    description: 'Search',
                    inputSchema: { type: 'object', properties: {} },
                  },
                ],
              }
            : {};
      await route.fulfill({
        status: 200,
        contentType: 'application/json',
        body: JSON.stringify({ jsonrpc: '2.0', id: body.id, result }),
      });
    });

    await page.evaluate(([k, v]) => localStorage.setItem(k, v), [AI_KEY, FAKE_KEY] as [
      string,
      string,
    ]);
    await page.reload();
    await waitForUniver(page);
    await page.getByTestId('panel-rail-ai').click();
    await expect(page.getByTestId('ai-panel')).toBeVisible();

    await page.getByTestId('ai-agent-toggle').click();
    await page.getByTestId('ai-mcp-add').click();
    await page.getByTestId('ai-mcp-input').fill('http://localhost/mcp-test/rpc');
    await page.keyboard.press('Enter');

    await expect(page.getByTestId('ai-mcp-section')).toContainText('1 tools', { timeout: 8000 });
  });
});
