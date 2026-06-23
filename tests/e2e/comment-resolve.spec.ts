import { expect, test } from '@playwright/test';
import { waitForUniver } from './_helpers';

/**
 * Comment resolve — Phase 3, T3.2 (what CommentsPanel now surfaces).
 *
 * The panel's resolve button dispatches `thread-comment.command.resolve-comment`
 * with `{ unitId, subUnitId, commentId, resolved: true }`. Resolving removes the
 * comment from the cell-location index, so it leaves the active list (open
 * threads only; a resolved view + reopen is a documented follow-up). This drives
 * that exact mechanism end-to-end — the menu-driven panel toggle is flaky in
 * headless and the panel's other actions aren't e2e'd, so we pin the command +
 * params the button relies on rather than the click.
 */
test('resolving a comment removes it from the active list', async ({ page }) => {
  test.setTimeout(60_000);
  await page.goto('/');
  await waitForUniver(page);

  // The threaded-comment plugin is lazy — ensure it's registered before use.
  await page.evaluate(() => window.__ensurePlugin__?.('threadComment'));

  const result = await page.evaluate(async () => {
    const api = window.__univerAPI!;
    const wb = api.getActiveWorkbook()!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const ws = wb.getActiveSheet() as any;
    await ws.getRange('A1').addComment({ dataStream: 'needs review\r\n' });

    const countAt = () => ws.getRange('A1').getComments?.()?.length ?? 0;
    const c0 = ws.getRange('A1').getComments?.()[0];
    const before = countAt();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const subUnitId = ws.getSheetId?.() ?? ws.getId?.();
    const cmdRes = await api.executeCommand('thread-comment.command.resolve-comment', {
      unitId: wb.getId(),
      subUnitId,
      commentId: c0?.id ?? c0?.getCommentData?.()?.id,
      resolved: true,
    });
    return { before, cmdRes, after: countAt() };
  });

  expect(result.before).toBe(1); // comment added + listed
  expect(result.cmdRes).toBe(true); // resolve command succeeded
  expect(result.after).toBe(0); // resolved → no longer in the active list
});
