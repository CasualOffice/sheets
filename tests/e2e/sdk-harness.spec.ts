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

/**
 * Exercises the SDK's `<CasualSheets>` editor directly via the dev-only
 * `/sdk-harness` route (apps/web/src/sdk-harness/SdkHarness.tsx). The app
 * normally renders its own `UniverSheet`, so this is the only coverage of the
 * published editor component. Verification surface for the SDK restructure.
 */

test.describe('SDK editor (CasualSheets) via /sdk-harness', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/sdk-harness');
    await page.waitForFunction(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      () => (window as any).__sdkHarnessReady === true,
      null,
      { timeout: 30_000 },
    );
  });

  test('boots and renders the grid (clean DI, no duplicate Univer)', async ({ page }) => {
    await expect(page.getByTestId('sdk-harness')).toBeVisible();
    // Univer renders the grid onto a sized <canvas> a frame or two after onReady;
    // a non-zero-size canvas means the render engine + plugin graph constructed
    // with no redi throw. (Don't use .first()/toBeVisible — the formula UI adds
    // 0-size overlay canvases that aren't the grid.)
    await page.waitForFunction(
      () => Array.from(document.querySelectorAll('canvas')).some((c) => c.clientWidth > 0),
      null,
      { timeout: 30_000 },
    );
  });

  test('formula engine computes (=1+2 → 3)', async ({ page }) => {
    const result = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const api = (window as any).__sdkHarnessAPI;
      // `api.univer` is the FUniver escape hatch on CasualSheetsAPI.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ws: any = api.univer.getActiveWorkbook().getActiveSheet();
      ws.getRange(0, 0).setValue({ f: '=1+2' });
      // Main-thread compute is near-synchronous, but poll to be safe.
      for (let i = 0; i < 100; i++) {
        const v = ws.getRange(0, 0).getValue();
        if (v === 3 || v === '3') return v;
        await new Promise((r) => setTimeout(r, 100));
      }
      return ws.getRange(0, 0).getValue();
    });
    expect(Number(result)).toBe(3);
  });

  test('trusted formula-result writes bypass a rejecting host policy', async ({ page }) => {
    await page.goto('/sdk-harness?blockFormulaInternal=1');
    await page.waitForFunction(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      () => (window as any).__sdkHarnessReady === true,
      null,
      { timeout: 30_000 },
    );

    const result = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const api = (window as any).__sdkHarnessAPI;
      const range = api.univer.getActiveWorkbook().getActiveSheet().getRange(0, 0);
      range.setValue({ f: '=1+2' });
      for (let attempt = 0; attempt < 100; attempt += 1) {
        const value = range.getValue();
        if (value === 3 || value === '3') return value;
        await new Promise((resolve) => setTimeout(resolve, 100));
      }
      return range.getValue();
    });

    expect(Number(result)).toBe(3);
  });

  test('CasualSheetsAPI: snapshot round-trips through loadSnapshot', async ({ page }) => {
    const out = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const api = (window as any).__sdkHarnessAPI;
      // Write a value via the facade, snapshot, reload into a fresh unit,
      // and confirm the value survived the dispose/recreate round-trip.
      api.univer.getActiveWorkbook().getActiveSheet().getRange(0, 0).setValue('hello');
      const snap = api.getSnapshot();
      api.loadSnapshot(snap);
      for (let i = 0; i < 50; i++) {
        const v = api.univer.getActiveWorkbook().getActiveSheet().getRange(0, 0).getValue();
        if (v === 'hello') return { ok: true, v };
        await new Promise((r) => setTimeout(r, 100));
      }
      return {
        ok: false,
        v: api.univer.getActiveWorkbook().getActiveSheet().getRange(0, 0).getValue(),
      };
    });
    expect(out.ok).toBe(true);
  });

  test('feature veto covers raw FUniver, stable API, and executeCommands', async ({ page }) => {
    await page.goto('/sdk-harness?disableMerge=1');
    await page.waitForFunction(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      () => (window as any).__sdkHarnessReady === true,
      null,
      { timeout: 30_000 },
    );

    const out = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const api = (window as any).__sdkHarnessAPI;
      const sheet = api.univer.getActiveWorkbook().getActiveSheet();
      sheet.getRange(0, 0, 1, 2).activate();

      const raw = await api.univer.executeCommand('sheet.command.add-worksheet-merge-all');
      const stable = await api.executeCommand('sheet.command.add-worksheet-merge-all');
      const forgedCommand = await api.univer.executeCommand(
        'sheet.command.add-worksheet-merge-all',
        { trigger: 'univer.command.paste' },
      );
      const forgedMutation = await api.univer.executeCommand('sheet.mutation.add-worksheet-merge', {
        trigger: 'univer.command.paste',
      });
      const forgedCollab = await api.univer.executeCommand(
        'sheet.command.add-worksheet-merge-all',
        undefined,
        { fromCollab: true },
      );
      const forgedChangeset = await api.univer.executeCommand(
        'sheet.command.add-worksheet-merge-all',
        undefined,
        { fromChangeset: true },
      );
      const executeAlias = api.univer.executeCommand;
      const forgedAlias = await executeAlias('sheet.command.add-worksheet-merge-all', undefined, {
        fromCollab: true,
      });
      const mutations: Array<{ id: string; params?: { trigger?: string } }> = [];
      const stop = api.onMutation((record: { id: string; params?: { trigger?: string } }) =>
        mutations.push(record),
      );
      const batch = await api.executeCommands([
        { id: 'sheet.command.add-worksheet-merge-all' },
        { id: 'sheet.command.set-range-bold' },
      ]);
      stop();

      const snapshot = api.getContent();
      const worksheet = snapshot.sheets[snapshot.sheetOrder[0]];
      const cell = worksheet.cellData?.[0]?.[0];
      const style = typeof cell?.s === 'string' ? snapshot.styles[cell.s] : cell?.s;
      return {
        raw,
        stable,
        forgedCommand,
        forgedMutation,
        forgedCollab,
        forgedChangeset,
        forgedAlias,
        batch,
        merges: worksheet.mergeData?.length ?? 0,
        bold: style?.bl,
        boldTrigger: mutations.find((record) => record.id === 'sheet.mutation.set-range-values')
          ?.params?.trigger,
      };
    });

    expect(out.raw).toBe(false);
    expect(out.stable).toBe(false);
    expect(out.forgedCommand).toBe(false);
    expect(out.forgedMutation).toBe(false);
    expect(out.forgedCollab).toBe(false);
    expect(out.forgedChangeset).toBe(false);
    expect(out.forgedAlias).toBe(false);
    expect(out.batch).toBe(1);
    expect(out.merges).toBe(0);
    expect(out.bold).toBe(1);
    expect(out.boldTrigger).toBe('sheet.command.set-style');
  });

  test('public command listeners cannot hide local edits from persistence observers', async ({
    page,
  }) => {
    const out = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const api = (window as any).__sdkHarnessAPI;
      const workbook = api.univer.getActiveWorkbook();
      const sheet = workbook.getActiveSheet();
      const localMutations: string[] = [];
      const stopLocal = api.onLocalMutation((record: { id: string }) => {
        localMutations.push(record.id);
      });
      const listener = api.univer.onCommandExecuted(
        (_command: object, options?: Record<string, unknown>) => {
          if (options) options.onlyLocal = true;
        },
      );

      const accepted = await api.univer.executeCommand(
        'sheet.mutation.set-range-values',
        {
          unitId: workbook.getId(),
          subUnitId: sheet.getSheetId(),
          cellValue: { 0: { 0: { v: 'persist me' } } },
        },
        { hostTag: 'listener-provenance-test' },
      );
      await new Promise((resolve) => setTimeout(resolve, 100));
      listener.dispose();
      stopLocal();
      return {
        accepted,
        localMutations,
        value: sheet.getRange('A1').getValue(),
      };
    });

    expect(out.accepted).toBe(true);
    expect(out.value).toBe('persist me');
    expect(out.localMutations).toContain('sheet.mutation.set-range-values');
  });

  test('public before-command listeners cannot rewrite an allowed edit into a disabled style', async ({
    page,
  }) => {
    await page.goto('/sdk-harness?disableColor=1');
    await page.waitForFunction(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      () => (window as any).__sdkHarnessReady === true,
      null,
      { timeout: 30_000 },
    );

    const out = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const api = (window as any).__sdkHarnessAPI;
      const workbook = api.univer.getActiveWorkbook();
      const sheet = workbook.getActiveSheet();
      const listener = api.univer.onBeforeCommandExecute(
        (command: {
          id: string;
          params?: { cellValue?: Record<number, Record<number, { s?: object }>> };
        }) => {
          if (command.id !== 'sheet.mutation.set-range-values') return;
          const cell = command.params?.cellValue?.[0]?.[0];
          if (cell) cell.s = { bg: { rgb: '#ff0000' } };
        },
      );
      const accepted = await api.univer.executeCommand('sheet.mutation.set-range-values', {
        unitId: workbook.getId(),
        subUnitId: sheet.getSheetId(),
        cellValue: { 0: { 0: { v: 'value only' } } },
      });
      listener.dispose();

      const snapshot = api.getContent();
      const worksheet = snapshot.sheets[snapshot.sheetOrder[0]];
      const cell = worksheet.cellData?.[0]?.[0];
      const style = typeof cell?.s === 'string' ? snapshot.styles[cell.s] : cell?.s;
      return { accepted, value: cell?.v, background: style?.bg };
    });

    expect(out.accepted).toBe(true);
    expect(out.value).toBe('value only');
    expect(out.background).toBeUndefined();
  });

  test('private bridge-style replay bypasses local policy and preserves convergence', async ({
    page,
  }) => {
    await page.goto('/sdk-harness?disableMerge=1');
    await page.waitForFunction(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      () => (window as any).__sdkHarnessReady === true,
      null,
      { timeout: 30_000 },
    );

    const out = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const w = window as any;
      const api = w.__sdkHarnessAPI;
      const workbook = api.univer.getActiveWorkbook();
      const sheet = workbook.getActiveSheet();
      const replay = await w.__sdkHarnessReplayCommand('sheet.mutation.add-worksheet-merge', {
        unitId: workbook.getId(),
        subUnitId: sheet.getSheetId(),
        ranges: [{ startRow: 4, startColumn: 0, endRow: 4, endColumn: 1 }],
      });
      const snapshot = api.getContent();
      const worksheet = snapshot.sheets[snapshot.sheetOrder[0]];
      return { replay, merges: worksheet.mergeData?.length ?? 0 };
    });

    expect(out).toEqual({ replay: { failures: 0 }, merges: 1 });
  });

  test('public provenance flags cannot hide a local edit from onLocalMutation', async ({
    page,
  }) => {
    const out = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const api = (window as any).__sdkHarnessAPI;
      const workbook = api.univer.getActiveWorkbook();
      const sheet = workbook.getActiveSheet();
      const mutations: Array<{ id: string; params?: object }> = [];
      const stop = api.onLocalMutation((record: { id: string; params?: object }) =>
        mutations.push(record),
      );

      const accepted = await api.univer.executeCommand(
        'sheet.mutation.set-range-values',
        {
          unitId: workbook.getId(),
          subUnitId: sheet.getSheetId(),
          cellValue: { 6: { 0: { v: 'record-me' } } },
        },
        {
          fromCollab: true,
          fromChangeset: true,
          fromFormula: true,
          applyFormulaCalculationResult: true,
          onlyLocal: true,
          syncOnly: true,
        },
      );
      stop();

      return {
        accepted,
        value: sheet.getRange(6, 0).getValue(),
        recorded: mutations.filter((record) => record.id === 'sheet.mutation.set-range-values')
          .length,
      };
    });

    expect(out).toEqual({ accepted: true, value: 'record-me', recorded: 1 });
  });

  test('local mutation stream excludes derived formula cache writes', async ({ page }) => {
    const out = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const api = (window as any).__sdkHarnessAPI;
      const sheet = api.univer.getActiveWorkbook().getActiveSheet();
      const mutations: Array<{ id: string; params?: { cellValue?: object } }> = [];
      const stop = api.onLocalMutation((record: { id: string; params?: { cellValue?: object } }) =>
        mutations.push(record),
      );

      sheet.getRange(7, 0).setValue({ f: '=1+2' });
      for (let index = 0; index < 50; index += 1) {
        if (Number(sheet.getRange(7, 0).getValue()) === 3) break;
        await new Promise((resolve) => setTimeout(resolve, 50));
      }
      stop();

      return {
        value: sheet.getRange(7, 0).getValue(),
        writes: mutations
          .filter((record) => record.id === 'sheet.mutation.set-range-values')
          .map((record) => record.params?.cellValue),
      };
    });

    expect(Number(out.value)).toBe(3);
    expect(out.writes).toHaveLength(1);
    expect(JSON.stringify(out.writes[0])).toContain('=1+2');
  });

  test('structural feature preflight blocks before a row mutation runs', async ({ page }) => {
    await page.goto('/sdk-harness?disableTables=1');
    await page.waitForFunction(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      () => (window as any).__sdkHarnessReady === true,
      null,
      { timeout: 30_000 },
    );

    const out = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const api = (window as any).__sdkHarnessAPI;
      const sheet = api.univer.getActiveWorkbook().getActiveSheet();
      const before = api.getContent();
      sheet.getRange(1, 0).activate();
      const accepted = await api.univer.executeCommand('sheet.command.insert-row-before', {
        value: 1,
      });
      const after = api.getContent();
      const beforeSheet = before.sheets[before.sheetOrder[0]];
      const afterSheet = after.sheets[after.sheetOrder[0]];
      return { accepted, beforeRows: beforeSheet.rowCount, afterRows: afterSheet.rowCount };
    });

    expect(out.accepted).toBe(false);
    expect(out.afterRows).toBe(out.beforeRows);
  });

  test('nested host policy errors fail closed without poisoning later commands', async ({
    page,
  }) => {
    await page.goto('/sdk-harness?throwOnMerge=1');
    await page.waitForFunction(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      () => (window as any).__sdkHarnessReady === true,
      null,
      { timeout: 30_000 },
    );

    const out = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const w = window as any;
      const api = w.__sdkHarnessAPI;
      const sheet = api.univer.getActiveWorkbook().getActiveSheet();
      sheet.getRange(8, 0, 1, 2).activate();
      const mutations: Array<{ id: string; params?: { trigger?: string } }> = [];
      const stop = api.onMutation((record: { id: string; params?: { trigger?: string } }) =>
        mutations.push(record),
      );

      const nested = await api.univer.executeCommand('test.command.nested-merge');
      sheet.getRange(8, 0).setValue('still-works');
      stop();

      const valueMutation = mutations.find(
        (record) => record.id === 'sheet.mutation.set-range-values',
      );
      return {
        nested,
        errors: w.__sdkHarnessErrors ?? [],
        value: sheet.getRange(8, 0).getValue(),
        staleParentTrigger: valueMutation?.params?.trigger === 'test.command.nested-merge',
      };
    });

    expect(out).toEqual({
      nested: false,
      errors: ['host policy exploded'],
      value: 'still-works',
      staleParentTrigger: false,
    });
  });

  test('throwing mutation observers cannot fail committed edits or later commands', async ({
    page,
  }) => {
    await page.goto('/sdk-harness?throwOnObserver=1');
    await page.waitForFunction(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      () => (window as any).__sdkHarnessReady === true,
      null,
      { timeout: 30_000 },
    );

    const out = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const w = window as any;
      const api = w.__sdkHarnessAPI;
      const workbook = api.univer.getActiveWorkbook();
      const sheet = workbook.getActiveSheet();
      const stop = api.onLocalMutation(() => {
        throw new Error('observer exploded');
      });
      const write = (row: number, value: string) =>
        api.univer.executeCommand('sheet.mutation.set-range-values', {
          unitId: workbook.getId(),
          subUnitId: sheet.getSheetId(),
          cellValue: { [row]: { 0: { v: value } } },
        });

      const first = await write(9, 'committed');
      stop();
      const second = await write(10, 'later');
      return {
        first,
        second,
        errors: w.__sdkHarnessErrors ?? [],
        firstValue: sheet.getRange(9, 0).getValue(),
        secondValue: sheet.getRange(10, 0).getValue(),
      };
    });

    expect(out).toEqual({
      first: true,
      second: true,
      errors: ['observer exploded'],
      firstValue: 'committed',
      secondValue: 'later',
    });
  });

  test('restricted compound paste is rejected before any nested mutation applies', async ({
    page,
  }) => {
    await page.goto('/sdk-harness?disableMerge=1');
    await page.waitForFunction(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      () => (window as any).__sdkHarnessReady === true,
      null,
      { timeout: 30_000 },
    );

    const out = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const api = (window as any).__sdkHarnessAPI;
      const sheet = api.univer.getActiveWorkbook().getActiveSheet();
      sheet.getRange(2, 0).setValue('keep');
      sheet.getRange(2, 0, 1, 2).activate();

      const accepted = await api.executeCommand('univer.command.paste', { value: 'default-paste' });
      // The underlying paste implementation is asynchronous even though its
      // command returns synchronously. Give an accidental handler invocation a
      // chance to mutate before inspecting the snapshot.
      await new Promise((resolve) => setTimeout(resolve, 100));

      const snapshot = api.getContent();
      const worksheet = snapshot.sheets[snapshot.sheetOrder[0]];
      return {
        accepted,
        first: worksheet.cellData?.[2]?.[0]?.v,
        second: worksheet.cellData?.[2]?.[1]?.v,
        merges: worksheet.mergeData?.length ?? 0,
      };
    });

    expect(out).toEqual({ accepted: false, first: 'keep', second: undefined, merges: 0 });
  });

  test('onChange streams a debounced snapshot after an edit', async ({ page }) => {
    const out = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const w = window as any;
      const before = w.__sdkHarnessChangeCount ?? 0;
      w.__sdkHarnessAPI.univer
        .getActiveWorkbook()
        .getActiveSheet()
        .getRange(5, 5)
        .setValue('changed');
      // Default debounce is 400ms; wait past it, then read the captured snapshot.
      for (let i = 0; i < 30; i++) {
        if ((w.__sdkHarnessChangeCount ?? 0) > before) break;
        await new Promise((r) => setTimeout(r, 100));
      }
      const snap = w.__sdkHarnessLastSnapshot;
      const cell = snap?.sheets?.[Object.keys(snap.sheets)[0]]?.cellData?.[5]?.[5];
      return {
        fired: (w.__sdkHarnessChangeCount ?? 0) > before,
        value: cell?.v,
      };
    });
    expect(out.fired).toBe(true);
    expect(out.value).toBe('changed');
  });

  test('lazy plugins idle-load (conditional-formatting command registers)', async ({ page }) => {
    // lazyPlugins defaults on; idleLoadAll registers the feature plugins after
    // first paint. The CF "add-conditional-rule" command is a stable marker that
    // the conditional-formatting plugin actually loaded into the editor.
    const has = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const hasCommand = (window as any).__sdkHarnessHasCommand as
        | ((id: string) => boolean)
        | undefined;
      for (let i = 0; i < 60; i++) {
        if (hasCommand?.('sheet.command.add-conditional-rule')) return true;
        await new Promise((r) => setTimeout(r, 100));
      }
      return hasCommand?.('sheet.command.add-conditional-rule') ?? false;
    });
    expect(has).toBe(true);
  });

  test('onBeforeCreateUnit lets a host register an extra plugin', async ({ page }) => {
    // crosshair-highlight is NOT in the SDK's plugin set — its command can only
    // appear if the onBeforeCreateUnit hook registered the plugin before the unit.
    await page.goto('/sdk-harness?beforeCreate=crosshair');
    await page.waitForFunction(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      () => (window as any).__sdkHarnessReady === true,
      null,
      { timeout: 30_000 },
    );
    const has = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const hasCommand = (window as any).__sdkHarnessHasCommand as
        | ((id: string) => boolean)
        | undefined;
      for (let i = 0; i < 30; i++) {
        if (hasCommand?.('sheet.operation.toggle-crosshair-highlight')) return true;
        await new Promise((r) => setTimeout(r, 100));
      }
      return hasCommand?.('sheet.operation.toggle-crosshair-highlight') ?? false;
    });
    expect(has).toBe(true);
  });

  test('without onBeforeCreateUnit, the extra plugin is absent', async ({ page }) => {
    // Control: the default mount must NOT have the crosshair command.
    const has = await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const hasCommand = (window as any).__sdkHarnessHasCommand as
        | ((id: string) => boolean)
        | undefined;
      return hasCommand?.('sheet.operation.toggle-crosshair-highlight') ?? false;
    });
    expect(has).toBe(false);
  });

  test('appearance="dark" flips Univer dark mode + container class', async ({ page }) => {
    await page.goto('/sdk-harness?appearance=dark');
    await page.waitForFunction(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      () => (window as any).__sdkHarnessReady === true,
      null,
      { timeout: 30_000 },
    );
    const dark = await page.evaluate(async () => {
      for (let i = 0; i < 30; i++) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        if ((window as any).__sdkHarnessIsDark?.()) break;
        await new Promise((r) => setTimeout(r, 100));
      }
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      return (window as any).__sdkHarnessIsDark?.();
    });
    expect(dark).toBe(true);
    // We mirror the class onto the editor container; Univer's Workbench also
    // applies it to <html> (its dark CSS is page-global by design).
    await expect(page.locator('[data-testid="casual-sheets"].univer-dark')).toHaveCount(1);
  });

  test('default appearance is light (no dark mode, no dark class)', async ({ page }) => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const dark = await page.evaluate(() => (window as any).__sdkHarnessIsDark?.());
    expect(dark).toBe(false);
    await expect(page.locator('.univer-dark')).toHaveCount(0);
  });

  test('CasualSheetsAPI: setTheme flips dark mode imperatively', async ({ page }) => {
    // Default mount is light; drive dark via the API ref.
    const result = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const w = window as any;
      const before = w.__sdkHarnessIsDark?.();
      w.__sdkHarnessAPI.setTheme('dark');
      for (let i = 0; i < 20; i++) {
        if (w.__sdkHarnessIsDark?.() === true) break;
        await new Promise((r) => setTimeout(r, 50));
      }
      const afterDark = w.__sdkHarnessIsDark?.();
      w.__sdkHarnessAPI.setTheme('light');
      const afterLight = w.__sdkHarnessIsDark?.();
      return { before, afterDark, afterLight };
    });
    expect(result.before).toBe(false);
    expect(result.afterDark).toBe(true);
    expect(result.afterLight).toBe(false);
  });

  test('chrome="minimal" renders the toolbar and Bold dispatches', async ({ page }) => {
    await page.goto('/sdk-harness?chrome=minimal');
    await page.waitForFunction(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      () => (window as any).__sdkHarnessReady === true,
      null,
      { timeout: 30_000 },
    );
    await expect(page.getByTestId('casual-sheets-toolbar')).toBeVisible();
    await expect(page.locator('[data-action="bold"]')).toBeVisible();
    // Put a value in A1, select it, then click Bold via the chrome toolbar.
    await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const api = (window as any).__sdkHarnessAPI;
      const ws = api.univer.getActiveWorkbook().getActiveSheet();
      ws.getRange(0, 0).setValue('x');
      ws.getRange(0, 0).activate();
      await new Promise((r) => setTimeout(r, 150));
    });
    await page.locator('[data-action="bold"]').click();
    // Verify via the snapshot: A1's resolved style has bold (bl === 1).
    const isBold = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const api = (window as any).__sdkHarnessAPI;
      for (let i = 0; i < 20; i++) {
        const snap = api.getSnapshot();
        const sheet = snap?.sheets?.[Object.keys(snap.sheets)[0]];
        const cell = sheet?.cellData?.[0]?.[0];
        // style is either inline (cell.s as object) or a ref into snap.styles.
        const style = cell && (typeof cell.s === 'string' ? snap.styles?.[cell.s] : cell.s);
        if (style?.bl === 1) return true;
        await new Promise((r) => setTimeout(r, 100));
      }
      return false;
    });
    expect(isBold).toBe(true);
  });

  test('chrome defaults to none (no toolbar/formula bar)', async ({ page }) => {
    await expect(page.getByTestId('casual-sheets-toolbar')).toHaveCount(0);
    await expect(page.getByTestId('casual-sheets-formula-bar')).toHaveCount(0);
  });

  test('chrome formula bar: name box tracks selection + edit commits', async ({ page }) => {
    await page.goto('/sdk-harness?chrome=minimal');
    await page.waitForFunction(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      () => (window as any).__sdkHarnessReady === true,
      null,
      { timeout: 30_000 },
    );
    await expect(page.getByTestId('casual-sheets-formula-bar')).toBeVisible();
    // Select B2 (row 1, col 1) — the name box should show "B2".
    await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const api = (window as any).__sdkHarnessAPI;
      api.univer.getActiveWorkbook().getActiveSheet().getRange(1, 1).activate();
      await new Promise((r) => setTimeout(r, 200));
    });
    await expect(page.getByTestId('cs-namebox-input')).toHaveValue('B2');
    // Type a formula into the bar and commit with Enter → B2 computes to 5.
    const input = page.getByTestId('casual-sheets-formula-input');
    await input.fill('=2+3');
    await input.press('Enter');
    const value = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const api = (window as any).__sdkHarnessAPI;
      for (let i = 0; i < 30; i++) {
        const v = api.univer.getActiveWorkbook().getActiveSheet().getRange(1, 1).getValue();
        if (v === 5 || v === '5') return v;
        await new Promise((r) => setTimeout(r, 100));
      }
      return api.univer.getActiveWorkbook().getActiveSheet().getRange(1, 1).getValue();
    });
    expect(Number(value)).toBe(5);
  });

  test('chrome status bar: selection stats (Average/Count/Sum)', async ({ page }) => {
    await page.goto('/sdk-harness?chrome=minimal');
    await page.waitForFunction(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      () => (window as any).__sdkHarnessReady === true,
      null,
      { timeout: 30_000 },
    );
    await expect(page.getByTestId('casual-sheets-status-bar')).toBeVisible();
    // Put 1,2,3 in A1:A3 and select the range → Sum 6, Count 3, Average 2.
    await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const api = (window as any).__sdkHarnessAPI;
      const ws = api.univer.getActiveWorkbook().getActiveSheet();
      ws.getRange(0, 0).setValue(1);
      ws.getRange(1, 0).setValue(2);
      ws.getRange(2, 0).setValue(3);
      ws.getRange('A1:A3').activate();
      await new Promise((r) => setTimeout(r, 250));
    });
    await expect(page.locator('[data-stat="sum"]')).toHaveText('Sum: 6');
    await expect(page.locator('[data-stat="count"]')).toHaveText('Count: 3');
    await expect(page.locator('[data-stat="average"]')).toHaveText('Average: 2');
    await expect(page.locator('[data-stat="num-count"]')).toHaveText('Numerical Count: 3');
    await expect(page.locator('[data-stat="min"]')).toHaveText('Min: 1');
    await expect(page.locator('[data-stat="max"]')).toHaveText('Max: 3');
  });

  test('viewing mode keeps selection, sheet navigation, and zoom usable', async ({ page }) => {
    await page.goto('/sdk-harness?chrome=minimal');
    await page.waitForFunction(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      () => (window as any).__sdkHarnessReady === true,
      null,
      { timeout: 30_000 },
    );

    const navigation = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const api = (window as any).__sdkHarnessAPI;
      const workbook = api.univer.getActiveWorkbook();
      const first = workbook.getActiveSheet();
      const second = workbook.insertSheet('Read-only navigation');
      workbook.setActiveSheet(first);
      api.setDocumentMode('viewing');
      workbook.setActiveSheet(second);
      second.getRange('B2').activate();
      await new Promise((resolve) => setTimeout(resolve, 100));
      return {
        mode: api.getDocumentMode(),
        activeSheetId: workbook.getActiveSheet().getSheetId(),
        expectedSheetId: second.getSheetId(),
        selection: api.getSelection(),
      };
    });

    expect(navigation.mode).toBe('viewing');
    expect(navigation.activeSheetId).toBe(navigation.expectedSheetId);
    expect(navigation.selection?.sheetId).toBe(navigation.expectedSheetId);
    expect(navigation.selection?.range).toMatchObject({
      startRow: 1,
      endRow: 1,
      startColumn: 1,
      endColumn: 1,
    });

    const level = page.getByTestId('cs-zoom-level');
    await expect(level).toHaveText('100%');
    await page.getByTestId('cs-zoom-in').click();
    await expect(level).toHaveText('110%');
  });

  test('chrome status bar: zoom control changes the zoom ratio', async ({ page }) => {
    await page.goto('/sdk-harness?chrome=minimal');
    await page.waitForFunction(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      () => (window as any).__sdkHarnessReady === true,
      null,
      { timeout: 30_000 },
    );
    // The zoom render module registers in sheets-ui's onRendered (after the grid
    // canvas paints) — wait for a sized canvas before driving zoom.
    await page.waitForFunction(
      () => Array.from(document.querySelectorAll('canvas')).some((c) => c.clientWidth > 0),
      null,
      { timeout: 30_000 },
    );
    const level = page.getByTestId('cs-zoom-level');
    await expect(level).toHaveText('100%');
    await page.getByTestId('cs-zoom-in').click();
    // Optimistic UI shows 110% immediately; the worksheet config confirms it.
    await expect(level).toHaveText('110%');
    const ratio = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const api = (window as any).__sdkHarnessAPI;
      for (let i = 0; i < 20; i++) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const z = (api.univer.getActiveWorkbook().getActiveSheet() as any)
          .getSheet()
          .getConfig().zoomRatio;
        if (Math.round(z * 100) === 110) return z;
        await new Promise((r) => setTimeout(r, 100));
      }
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      return (api.univer.getActiveWorkbook().getActiveSheet() as any).getSheet().getConfig()
        .zoomRatio;
    });
    expect(Math.round(ratio * 100)).toBe(110);
    // Click the level to reset to 100%.
    await level.click();
    await expect(level).toHaveText('100%');
  });

  test('chrome status bar: Count includes text cells, numeric stats skip them', async ({
    page,
  }) => {
    await page.goto('/sdk-harness?chrome=minimal');
    await page.waitForFunction(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      () => (window as any).__sdkHarnessReady === true,
      null,
      { timeout: 30_000 },
    );
    // A1=10, A2="x", A3=20 → Count 3 (non-empty), Numerical Count 2, Sum 30.
    await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const api = (window as any).__sdkHarnessAPI;
      const ws = api.univer.getActiveWorkbook().getActiveSheet();
      ws.getRange(0, 0).setValue(10);
      ws.getRange(1, 0).setValue('x');
      ws.getRange(2, 0).setValue(20);
      ws.getRange('A1:A3').activate();
      await new Promise((r) => setTimeout(r, 250));
    });
    await expect(page.locator('[data-stat="count"]')).toHaveText('Count: 3');
    await expect(page.locator('[data-stat="num-count"]')).toHaveText('Numerical Count: 2');
    await expect(page.locator('[data-stat="sum"]')).toHaveText('Sum: 30');
  });

  test('chrome toolbar: reflects active cell (Bold active state + font size)', async ({ page }) => {
    await page.goto('/sdk-harness?chrome=minimal');
    await page.waitForFunction(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      () => (window as any).__sdkHarnessReady === true,
      null,
      { timeout: 30_000 },
    );
    // Select A1 — bold not active yet.
    await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const api = (window as any).__sdkHarnessAPI;
      api.univer.getActiveWorkbook().getActiveSheet().getRange(0, 0).activate();
      await new Promise((r) => setTimeout(r, 200));
    });
    await expect(page.locator('[data-action="bold"]')).not.toHaveAttribute('data-active', 'true');
    // Bold it via the toolbar → the button reflects the active state.
    await page.locator('[data-action="bold"]').click();
    await expect(page.locator('[data-action="bold"]')).toHaveAttribute('data-active', 'true');
  });

  test('chrome toolbar: font size dropdown applies to the cell', async ({ page }) => {
    await page.goto('/sdk-harness?chrome=minimal');
    await page.waitForFunction(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      () => (window as any).__sdkHarnessReady === true,
      null,
      { timeout: 30_000 },
    );
    await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const api = (window as any).__sdkHarnessAPI;
      const ws = api.univer.getActiveWorkbook().getActiveSheet();
      ws.getRange(0, 0).setValue('x');
      ws.getRange(0, 0).activate();
      await new Promise((r) => setTimeout(r, 150));
    });
    await page.getByTestId('cs-font-size').selectOption('24');
    const fs = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const api = (window as any).__sdkHarnessAPI;
      for (let i = 0; i < 20; i++) {
        const snap = api.getSnapshot();
        const sheet = snap?.sheets?.[Object.keys(snap.sheets)[0]];
        const cell = sheet?.cellData?.[0]?.[0];
        const style = cell && (typeof cell.s === 'string' ? snap.styles?.[cell.s] : cell.s);
        if (style?.fs === 24) return 24;
        await new Promise((r) => setTimeout(r, 100));
      }
      return null;
    });
    expect(fs).toBe(24);
  });

  test('chrome toolbar: Merge cells merges the selection', async ({ page }) => {
    await page.goto('/sdk-harness?chrome=minimal');
    await page.waitForFunction(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      () => (window as any).__sdkHarnessReady === true,
      null,
      { timeout: 30_000 },
    );
    await expect(page.locator('[data-action="merge"]')).toBeVisible();
    await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const api = (window as any).__sdkHarnessAPI;
      api.univer.getActiveWorkbook().getActiveSheet().getRange('A1:B2').activate();
      await new Promise((r) => setTimeout(r, 150));
    });
    await page.locator('[data-action="merge"]').click();
    const merges = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const api = (window as any).__sdkHarnessAPI;
      for (let i = 0; i < 20; i++) {
        const snap = api.getSnapshot();
        const sheet = snap?.sheets?.[Object.keys(snap.sheets)[0]];
        if ((sheet?.mergeData?.length ?? 0) > 0) return sheet.mergeData.length;
        await new Promise((r) => setTimeout(r, 100));
      }
      return 0;
    });
    expect(merges).toBeGreaterThan(0);
  });

  test('chrome flips to dark with appearance="dark"', async ({ page }) => {
    await page.goto('/sdk-harness?chrome=minimal&appearance=dark');
    await page.waitForFunction(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      () => (window as any).__sdkHarnessReady === true,
      null,
      { timeout: 30_000 },
    );
    const bg = await page
      .getByTestId('casual-sheets-toolbar')
      .evaluate((el) => getComputedStyle(el).backgroundColor);
    // Dark chrome bg is the design-system surface-strip #2a2e35 → rgb(42, 46, 53).
    expect(bg).toBe('rgb(42, 46, 53)');
  });

  test('chrome formula bar: function autocomplete completes', async ({ page }) => {
    await page.goto('/sdk-harness?chrome=minimal');
    await page.waitForFunction(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      () => (window as any).__sdkHarnessReady === true,
      null,
      { timeout: 30_000 },
    );
    const input = page.getByTestId('casual-sheets-formula-input');
    const suggestions = page.getByTestId('cs-formula-suggestions');
    // The formula input is disabled until the SDK api binds; don't type early.
    await expect(input).toBeEnabled();
    // Right after the harness signals ready, background idle-plugin-load command
    // bursts re-render the chrome and can swallow the first keystrokes' draft
    // before the autocomplete renders. Retry the type→suggestions step (rather
    // than racing a single fill) until the dropdown is actually up — generous
    // budgets so a CPU-starved run under full e2e load still converges.
    await expect(async () => {
      await input.click();
      await input.clear();
      await input.pressSequentially('=SU');
      await expect(suggestions).toBeVisible({ timeout: 5_000 });
    }).toPass({ timeout: 30_000 });
    await expect(page.getByTestId('cs-formula-suggestion-SUM')).toBeVisible();
    // Complete via keyboard (ArrowDown to SUM, then Enter). Keyboard-driven so
    // it's deterministic under suite load — clicking the item is flaky while the
    // chrome re-renders from background idle-plugin-load command bursts.
    // Suggestions for "=SU" are [SUBSTITUTE, SUM, …]; ArrowDown once → SUM.
    await input.press('ArrowDown');
    await input.press('Enter');
    await expect(input).toHaveValue('=SUM(');
  });

  test('chrome toolbar: wrap text applies', async ({ page }) => {
    await page.goto('/sdk-harness?chrome=minimal');
    await page.waitForFunction(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      () => (window as any).__sdkHarnessReady === true,
      null,
      { timeout: 30_000 },
    );
    await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const api = (window as any).__sdkHarnessAPI;
      api.univer.getActiveWorkbook().getActiveSheet().getRange(0, 0).setValue('x');
      api.univer.getActiveWorkbook().getActiveSheet().getRange(0, 0).activate();
      await new Promise((r) => setTimeout(r, 150));
    });
    await page.locator('[data-action="wrap-text"]').click();
    const wrapped = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const api = (window as any).__sdkHarnessAPI;
      for (let i = 0; i < 20; i++) {
        const snap = api.getSnapshot();
        const sheet = snap?.sheets?.[Object.keys(snap.sheets)[0]];
        const cell = sheet?.cellData?.[0]?.[0];
        const style = cell && (typeof cell.s === 'string' ? snap.styles?.[cell.s] : cell.s);
        if (style?.tb === 3) return true;
        await new Promise((r) => setTimeout(r, 100));
      }
      return false;
    });
    expect(wrapped).toBe(true);
  });

  test('chrome menu bar: View menu renders (freeze)', async ({ page }) => {
    await page.goto('/sdk-harness?chrome=minimal');
    await page.waitForFunction(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      () => (window as any).__sdkHarnessReady === true,
      null,
      { timeout: 30_000 },
    );
    await page.locator('[data-menu="view"]').click();
    await expect(page.getByTestId('cs-menuitem-freeze-row')).toBeVisible();
  });

  test('chrome color picker: text color applies + popover closes', async ({ page }) => {
    await page.goto('/sdk-harness?chrome=minimal');
    await page.waitForFunction(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      () => (window as any).__sdkHarnessReady === true,
      null,
      { timeout: 30_000 },
    );
    await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const api = (window as any).__sdkHarnessAPI;
      const ws = api.univer.getActiveWorkbook().getActiveSheet();
      ws.getRange(0, 0).setValue('x');
      ws.getRange(0, 0).activate();
      await new Promise((r) => setTimeout(r, 150));
    });
    await page.locator('[data-testid="cs-color-text"]').click();
    await expect(page.getByTestId('cs-color-popover')).toBeVisible();
    await page.locator('[data-color="#0e7490"]').click();
    await expect(page.getByTestId('cs-color-popover')).toHaveCount(0);
    const colored = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const api = (window as any).__sdkHarnessAPI;
      for (let i = 0; i < 20; i++) {
        const snap = api.getSnapshot();
        const sheet = snap?.sheets?.[Object.keys(snap.sheets)[0]];
        const cell = sheet?.cellData?.[0]?.[0];
        const style = cell && (typeof cell.s === 'string' ? snap.styles?.[cell.s] : cell.s);
        const rgb = style?.cl?.rgb ?? style?.cl;
        if (typeof rgb === 'string' && rgb.toLowerCase() === '#0e7490') return true;
        await new Promise((r) => setTimeout(r, 100));
      }
      return false;
    });
    expect(colored).toBe(true);
  });

  test('chrome menu bar: Format → Bold dispatches', async ({ page }) => {
    await page.goto('/sdk-harness?chrome=minimal');
    await page.waitForFunction(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      () => (window as any).__sdkHarnessReady === true,
      null,
      { timeout: 30_000 },
    );
    await expect(page.getByTestId('cs-menubar')).toBeVisible();
    await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const api = (window as any).__sdkHarnessAPI;
      api.univer.getActiveWorkbook().getActiveSheet().getRange(0, 0).setValue('x');
      api.univer.getActiveWorkbook().getActiveSheet().getRange(0, 0).activate();
      await new Promise((r) => setTimeout(r, 150));
    });
    await page.locator('[data-menu="format"]').click();
    await expect(page.getByTestId('cs-menuitem-bold')).toBeVisible();
    await page.getByTestId('cs-menuitem-bold').click();
    const bold = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const api = (window as any).__sdkHarnessAPI;
      for (let i = 0; i < 20; i++) {
        const snap = api.getSnapshot();
        const sheet = snap?.sheets?.[Object.keys(snap.sheets)[0]];
        const cell = sheet?.cellData?.[0]?.[0];
        const style = cell && (typeof cell.s === 'string' ? snap.styles?.[cell.s] : cell.s);
        if (style?.bl === 1) return true;
        await new Promise((r) => setTimeout(r, 100));
      }
      return false;
    });
    expect(bold).toBe(true);
  });

  test('chrome name box: typing a ref navigates the selection', async ({ page }) => {
    await page.goto('/sdk-harness?chrome=minimal');
    await page.waitForFunction(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      () => (window as any).__sdkHarnessReady === true,
      null,
      { timeout: 30_000 },
    );
    const input = page.getByTestId('cs-namebox-input');
    await expect(input).toBeVisible();
    // Re-issue fill+Enter inside the poll: under CI load the first Enter can land
    // before the name-box→selection wiring is live, leaving the selection at A1.
    // Retrying the whole interaction self-heals that (idempotent — re-navigating
    // to C5 is a no-op once it's selected) instead of one-shot-then-hope.
    await expect
      .poll(
        async () => {
          await input.click();
          await input.fill('C5');
          await input.press('Enter');
          return page.evaluate(() => {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const s = (window as any).__sdkHarnessAPI?.getSelection?.();
            return s ? { startRow: s.range.startRow, startColumn: s.range.startColumn } : null;
          });
        },
        { timeout: 15_000, intervals: [200, 400, 600, 800] },
      )
      .toEqual({ startRow: 4, startColumn: 2 });
  });

  test('chrome sheet tabs: add a sheet → second tab appears and activates', async ({ page }) => {
    await page.goto('/sdk-harness?chrome=minimal');
    await page.waitForFunction(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      () => (window as any).__sdkHarnessReady === true,
      null,
      { timeout: 30_000 },
    );
    const strip = page.getByTestId('casual-sheets-tabs');
    await expect(strip).toBeVisible();
    // The default workbook has exactly one sheet tab.
    await expect(strip.getByRole('tab')).toHaveCount(1);
    await page.getByTestId('cs-tab-add').click();
    // A second tab should appear, and the new one becomes active.
    await expect(strip.getByRole('tab')).toHaveCount(2);
    const sheetCount = await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const api = (window as any).__sdkHarnessAPI;
      return api.univer.getActiveWorkbook().getSheets().length;
    });
    expect(sheetCount).toBe(2);
  });

  test('chrome sheet tabs: double-click rename commits the new name', async ({ page }) => {
    await page.goto('/sdk-harness?chrome=minimal');
    await page.waitForFunction(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      () => (window as any).__sdkHarnessReady === true,
      null,
      { timeout: 30_000 },
    );
    const tab = page.getByTestId('casual-sheets-tabs').getByRole('tab').first();
    const input = page.getByTestId('cs-tab-rename-input');
    // Background idle-plugin-load command bursts re-render the chrome just after
    // the harness signals ready and can swallow the first double-click before
    // the rename input mounts. Retry the dblclick→input step (rather than racing
    // a single dblclick) until the inline editor is actually up.
    await expect(async () => {
      await tab.dblclick();
      await expect(input).toBeVisible({ timeout: 2_000 });
    }).toPass({ timeout: 30_000 });
    await input.fill('Budget');
    await input.press('Enter');
    const name = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const api = (window as any).__sdkHarnessAPI;
      for (let i = 0; i < 20; i++) {
        const n = api.univer.getActiveWorkbook().getActiveSheet().getSheetName();
        if (n === 'Budget') return n;
        await new Promise((r) => setTimeout(r, 100));
      }
      return api.univer.getActiveWorkbook().getActiveSheet().getSheetName();
    });
    expect(name).toBe('Budget');
  });

  test('chrome toolbar: borders dropdown applies a border to the selection', async ({ page }) => {
    await page.goto('/sdk-harness?chrome=minimal');
    await page.waitForFunction(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      () => (window as any).__sdkHarnessReady === true,
      null,
      { timeout: 30_000 },
    );
    // No border styles initially.
    const before = await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const snap = (window as any).__sdkHarnessAPI.getSnapshot();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      return Object.values(snap.styles || {}).filter((s: any) => s && s.bd).length;
    });
    expect(before).toBe(0);
    // Open the borders dropdown and apply "All borders".
    await page.getByTestId('cs-borders-button').click();
    await expect(page.getByTestId('cs-borders-popover')).toBeVisible();
    await page.getByTestId('cs-border-all').click();
    // The popover closes and a border style now exists.
    await expect(page.getByTestId('cs-borders-popover')).toHaveCount(0);
    const after = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const api = (window as any).__sdkHarnessAPI;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const hasBorder = () => Object.values(api.getSnapshot().styles || {}).some((s: any) => s?.bd);
      for (let i = 0; i < 20; i++) {
        if (hasBorder()) return true;
        await new Promise((r) => setTimeout(r, 100));
      }
      return false;
    });
    expect(after).toBe(true);
  });

  test('chrome toolbar: AutoSum inserts =SUM below the selection', async ({ page }) => {
    await page.goto('/sdk-harness?chrome=minimal');
    await page.waitForFunction(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      () => (window as any).__sdkHarnessReady === true,
      null,
      { timeout: 30_000 },
    );
    // 1,2,3 in A1:A3, select the range → AutoSum → Sum drops =SUM(A1:A3) in A4.
    await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const api = (window as any).__sdkHarnessAPI;
      const ws = api.univer.getActiveWorkbook().getActiveSheet();
      ws.getRange(0, 0).setValue(1);
      ws.getRange(1, 0).setValue(2);
      ws.getRange(2, 0).setValue(3);
      ws.getRange('A1:A3').activate();
      await new Promise((r) => setTimeout(r, 200));
    });
    await page.getByTestId('cs-autosum-button').click();
    await expect(page.getByTestId('cs-autosum-popover')).toBeVisible();
    await page.getByTestId('cs-autosum-SUM').click();
    await expect(page.getByTestId('cs-autosum-popover')).toHaveCount(0);
    const out = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const api = (window as any).__sdkHarnessAPI;
      const ws = api.univer.getActiveWorkbook().getActiveSheet();
      for (let i = 0; i < 30; i++) {
        const f = ws.getRange(3, 0).getFormula?.();
        const v = ws.getRange(3, 0).getValue();
        if (v === 6 || v === '6') return { f, v };
        await new Promise((r) => setTimeout(r, 100));
      }
      return { f: ws.getRange(3, 0).getFormula?.(), v: ws.getRange(3, 0).getValue() };
    });
    expect(out.f).toBe('=SUM(A1:A3)');
    expect(Number(out.v)).toBe(6);
  });

  test('onSave fires on Ctrl/Cmd+S with the snapshot', async ({ page }) => {
    await page.goto('/sdk-harness?chrome=minimal');
    await page.waitForFunction(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      () => (window as any).__sdkHarnessReady === true,
      null,
      { timeout: 30_000 },
    );
    // Focus inside the editor so the capture-phase Ctrl+S handler sees the key.
    await page.getByTestId('casual-sheets-formula-input').click();
    await page.keyboard.press('Control+s');
    const out = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const w = window as any;
      for (let i = 0; i < 20; i++) {
        if ((w.__sdkHarnessSaveCount ?? 0) > 0) break;
        await new Promise((r) => setTimeout(r, 50));
      }
      return {
        count: w.__sdkHarnessSaveCount ?? 0,
        hasSnapshot: !!w.__sdkHarnessLastSaved?.sheets,
      };
    });
    expect(out.count).toBeGreaterThan(0);
    expect(out.hasSnapshot).toBe(true);
  });

  test('CasualSheetsAPI: getSelection returns the active range', async ({ page }) => {
    const sel = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const api = (window as any).__sdkHarnessAPI;
      api.univer.getActiveWorkbook().getActiveSheet().getRange(2, 3).activate();
      // Selection commands settle on the next frame.
      await new Promise((r) => setTimeout(r, 200));
      return api.getSelection();
    });
    expect(sel).not.toBeNull();
    expect(sel.range.startRow).toBe(2);
    expect(sel.range.startColumn).toBe(3);
    expect(typeof sel.sheetId).toBe('string');
  });
});
