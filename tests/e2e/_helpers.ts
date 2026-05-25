import type { Page } from '@playwright/test';

/**
 * Returns the main grid canvas (the one users interact with).
 * Univer also renders a hidden formula-editor canvas; that one has no stable
 * id and we never want to target it from tests.
 */
export function mainCanvas(page: Page) {
  // Canvas id encodes the workbook unit id, which is now allocated dynamically
  // per workbook (so Open / New don't collide on the previous id).
  return page.locator('[id^="univer-sheet-main-canvas_"]');
}

/**
 * Wait for Univer to finish mounting + dismiss the home screen by
 * default. Tests that explicitly need the home gallery (the
 * home-screen.spec suite) pass `{ keepHome: true }` to opt out.
 */
export async function waitForUniver(page: Page, opts: { keepHome?: boolean } = {}) {
  await mainCanvas(page).waitFor({ timeout: 15_000 });
  await page.waitForFunction(() => Boolean(window.__univerAPI), null, { timeout: 5_000 });
  // The home-screen template gallery (added in v0.1.0) is a full-viewport
  // overlay above the editor. It blocks every pointer-event-driven test.
  // Tests assume the editor is interactable straight after waitForUniver,
  // so dismiss the home screen here. Production users dismiss it via the
  // close-X or by picking a template; tests just need it gone.
  if (!opts.keepHome) {
    await dismissHomeScreen(page);
  }
}

/**
 * Click the home-screen close button if it's visible. No-op when home
 * is already dismissed (autosave-restore flow, collab URL, etc.).
 * Tolerates absence — older builds without the home gallery still work.
 *
 * Race-avoidance: the app has in-flight useEffects that dismiss home
 * on their own when there's an autosave record or a collab URL. Yield
 * for ~250 ms before deciding whether to dismiss manually. Without
 * this, we click the close button BEFORE the autosave-driven useEffect
 * fires, which (a) registers a click → autosave-driver flips
 * `userInteracted` true → the empty Untitled workbook overwrites the
 * seeded autosave record 5 s later, and (b) the banner mounts
 * post-dismiss but reads an empty record and never appears. The grace
 * period lets the app dismiss home itself when it's going to.
 *
 * Dismissing via Page.evaluate (setting `__casualE2E_dismissHome`)
 * could replace the click, but the current click path is exercised in
 * production and worth keeping covered.
 */
async function dismissHomeScreen(page: Page) {
  try {
    const home = page.getByTestId('home-screen');
    if ((await home.count()) === 0) return;
    // Grace period — see comment above.
    await page.waitForTimeout(250);
    if (!(await home.isVisible({ timeout: 200 }).catch(() => false))) return;
    const closeBtn = home.getByRole('button', { name: /close|dismiss/i }).first();
    if ((await closeBtn.count()) > 0) {
      await closeBtn.click({ timeout: 2_000 }).catch(() => undefined);
    } else {
      // Fallback: dispatch keydown Escape on the home overlay.
      await page.keyboard.press('Escape').catch(() => undefined);
    }
    await home.waitFor({ state: 'hidden', timeout: 2_000 }).catch(() => undefined);
  } catch {
    // If anything throws, leave the home up — the test will fail with a
    // clear locator error rather than hanging silently.
  }
}

export async function selectRange(page: Page, a1: string) {
  await page.evaluate((cell) => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    ws.getRange(cell).activate();
  }, a1);
}

export async function readCell(page: Page, a1: string) {
  return page.evaluate((cell) => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws: any = api.getActiveWorkbook()!.getActiveSheet();
    return ws.getRange(cell).getCellData();
  }, a1);
}

/**
 * Reads the resolved IStyleData for a cell — handles both inline styles
 * (cellData.s is the object) and interned styles (cellData.s is an id into
 * the workbook's style table).
 */
export async function readStyle(page: Page, a1: string) {
  return page.evaluate((cell) => {
    const api = window.__univerAPI!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const wb: any = api.getActiveWorkbook()!;
    const ws = wb.getActiveSheet();
    const data = ws.getRange(cell).getCellData();
    if (!data) return null;
    if (typeof data.s === 'string') {
      return wb.getWorkbook().getStyles().get(data.s) ?? null;
    }
    return data.s ?? null;
  }, a1);
}
