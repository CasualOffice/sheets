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

import { strict as assert } from 'node:assert';
import { test } from 'node:test';

import {
  createCommandPolicy,
  disabledFeatureForCommand,
  type BeforeCommandContext,
  type BeforeCommandPolicy,
} from './command-policy';
import { runSteps } from './scripting';

test('only explicitly-disabled, known features veto commands', () => {
  const merge = { id: 'sheet.command.add-worksheet-merge-all' };
  assert.equal(disabledFeatureForCommand(merge, undefined), undefined);
  assert.equal(disabledFeatureForCommand(merge, {}), undefined);
  assert.equal(disabledFeatureForCommand(merge, { merge: true }), undefined);
  assert.equal(disabledFeatureForCommand(merge, { futureCapability: false }), undefined);
  assert.equal(disabledFeatureForCommand(merge, { merge: false }), 'merge');
});

test('command-backed chrome features map every engine entry path', () => {
  const cases: Array<[string, BeforeCommandContext]> = [
    ['history', { id: 'univer.command.undo' }],
    ['clipboard', { id: 'sheet.command.paste-formula' }],
    ['format-painter', { id: 'sheet.command.apply-format-painter' }],
    ['font', { id: 'sheet.command.set-range-font-family' }],
    ['font-style', { id: 'sheet.command.set-range-bold' }],
    ['color', { id: 'sheet.command.set-range-text-color' }],
    ['borders', { id: 'sheet.command.set-border-basic' }],
    ['alignment', { id: 'sheet.command.set-horizontal-text-align' }],
    ['merge', { id: 'sheet.mutation.remove-worksheet-merge' }],
    ['number', { id: 'sheet.mutation.set.numfmt' }],
    ['clear-format', { id: 'sheet.command.clear-selection-format' }],
    ['tables', { id: 'sheet.command.table-insert-row' }],
    ['conditionalFormatting', { id: 'sheet.command.add-color-scale-conditional-rule' }],
    ['filter', { id: 'sheet.command.smart-toggle-filter' }],
    ['dataValidation', { id: 'sheet.command.addDataValidation' }],
  ];

  for (const [feature, command] of cases) {
    assert.equal(
      disabledFeatureForCommand(command, { [feature]: false }),
      feature,
      `expected ${command.id} to be governed by ${feature}`,
    );
  }
});

test('facade style commands and recorded mutation replays cannot bypass feature flags', () => {
  // FRange.setFontWeight() dispatches the generic set-style command.
  assert.equal(
    disabledFeatureForCommand(
      { id: 'sheet.command.set-style', params: { style: { type: 'bl', value: 1 } } },
      { 'font-style': false },
    ),
    'font-style',
  );

  // onMutation() records the derived set-range-values mutation. executeCommands
  // replays this shape rather than the toolbar's original command id.
  assert.equal(
    disabledFeatureForCommand(
      {
        id: 'sheet.mutation.set-range-values',
        params: {
          trigger: 'sheet.command.set-style',
          cellValue: { 0: { 0: { s: { bg: { rgb: '#ff0000' } } } } },
        },
      },
      { color: false },
    ),
    'color',
  );

  // Number format is carried as the `n` style key through the same generic
  // command/mutation paths; it must not depend on a dedicated numfmt id.
  assert.equal(
    disabledFeatureForCommand(
      { id: 'sheet.command.set-style', params: { style: { type: 'n', value: '#,##0' } } },
      { number: false },
    ),
    'number',
  );
  assert.equal(
    disabledFeatureForCommand(
      {
        id: 'sheet.mutation.set-range-values',
        params: { cellValue: { 0: { 0: { s: { n: { pattern: '#,##0' } } } } } },
      },
      { number: false },
    ),
    'number',
  );
});

test('all Univer 0.25 style keys map to their formatting capability', () => {
  const cases: Array<[string, string]> = [
    ['va', 'font-style'],
    ['bbl', 'font-style'],
    ['td', 'alignment'],
    ['pd', 'alignment'],
  ];

  for (const [styleKey, feature] of cases) {
    assert.equal(
      disabledFeatureForCommand(
        {
          id: 'sheet.mutation.set-range-values',
          params: { cellValue: { 0: { 0: { s: { [styleKey]: {} } } } } },
        },
        { [feature]: false },
      ),
      feature,
      `expected style key ${styleKey} to be governed by ${feature}`,
    );
  }
});

test('opaque style ids and style-clearing mutations fail closed', () => {
  assert.equal(
    disabledFeatureForCommand(
      {
        id: 'sheet.mutation.set-range-values',
        params: { cellValue: { 0: { 0: { s: 'workbook-style-7' } } } },
      },
      { color: false },
    ),
    'color',
  );
  assert.equal(
    disabledFeatureForCommand(
      {
        id: 'sheet.mutation.set-range-values',
        params: { cellValue: { 0: { 0: { s: null } } } },
      },
      { font: false },
    ),
    'font',
  );
  assert.equal(
    disabledFeatureForCommand(
      {
        id: 'sheet.mutation.set-range-values',
        params: { cellValue: { 0: { 0: null } } },
      },
      { alignment: false },
    ),
    'alignment',
  );
  assert.equal(
    disabledFeatureForCommand(
      { id: 'sheet.command.set-style', params: { style: { type: 'future-style' } } },
      { 'font-style': false },
    ),
    'font-style',
  );
  assert.equal(
    disabledFeatureForCommand(
      {
        id: 'sheet.mutation.set-range-values',
        params: { cellValue: { 0: { 0: { s: { futureStyleKey: 1 } } } } },
      },
      { alignment: false },
    ),
    'alignment',
  );
});

test('style inspection is skipped when no formatter is disabled and bounded otherwise', () => {
  let cellReads = 0;
  const row: Record<string, unknown> = {};
  for (let index = 0; index < 15_000; index += 1) {
    Object.defineProperty(row, String(index), {
      enumerable: true,
      get() {
        cellReads += 1;
        return { v: index };
      },
    });
  }
  const command = {
    id: 'sheet.mutation.set-range-values',
    params: { cellValue: { 0: row } },
  };

  assert.equal(disabledFeatureForCommand(command, { merge: false }), undefined);
  assert.equal(cellReads, 0, 'a non-formatting restriction must not inspect cell styles');

  assert.equal(disabledFeatureForCommand(command, { font: false }), 'font');
  assert.ok(cellReads <= 10_000, `style inspection read ${cellReads} cells`);
});

test('compound edits are vetoed before a disabled nested mutation can partially apply', () => {
  // Default paste may apply values/styles before it reaches merge mutations.
  // Blocking the parent is the only atomic outcome under Univer 0.25.
  assert.equal(
    disabledFeatureForCommand({ id: 'univer.command.paste' }, { merge: false }),
    'merge',
  );
  assert.equal(
    disabledFeatureForCommand(
      { id: 'sheet.command.paste-by-short-key', params: { textContent: 'A\tB' } },
      { color: false },
    ),
    'color',
  );

  // Explicit value-only paste has no style or merge mutations and remains a
  // useful safe path in a restricted editor.
  assert.equal(
    disabledFeatureForCommand(
      { id: 'univer.command.paste', params: { value: 'special-paste-value' } },
      { merge: false, color: false },
    ),
    undefined,
  );
  assert.equal(
    disabledFeatureForCommand(
      { id: 'univer.command.paste', params: { type: 'PASTE_VALUE' } },
      { merge: false, color: false },
    ),
    undefined,
  );

  // Clear-all and clear-format both clear every style property through a
  // generic set-range-values mutation.
  assert.equal(
    disabledFeatureForCommand({ id: 'sheet.command.clear-selection-all' }, { font: false }),
    'font',
  );
  assert.equal(
    disabledFeatureForCommand(
      { id: 'sheet.command.clear-selection-all' },
      { 'clear-format': false },
    ),
    'clear-format',
  );

  // Clear and format-painter handlers also append merge/CF/DV mutations. A
  // child veto would be too late after the leading style mutation.
  assert.equal(
    disabledFeatureForCommand(
      { id: 'sheet.command.clear-selection-format' },
      { conditionalFormatting: false },
    ),
    'conditionalFormatting',
  );
  assert.equal(
    disabledFeatureForCommand(
      { id: 'sheet.command.clear-selection-all' },
      { dataValidation: false },
    ),
    'dataValidation',
  );
  assert.equal(
    disabledFeatureForCommand({ id: 'sheet.command.apply-format-painter' }, { merge: false }),
    'merge',
  );
  assert.equal(
    disabledFeatureForCommand({ id: 'sheet.command.remove-worksheet-merge' }, { color: false }),
    'color',
  );
});

test('paste preflight accounts for each special mode before trusting its children', () => {
  const paste = (value: string, features: Record<string, boolean>) =>
    disabledFeatureForCommand({ id: 'univer.command.paste', params: { value } }, features);

  assert.equal(paste('default-paste', { conditionalFormatting: false }), 'conditionalFormatting');
  assert.equal(paste('default-paste', { dataValidation: false }), 'dataValidation');
  assert.equal(paste('special-paste-format', { merge: false }), 'merge');
  assert.equal(
    paste('special-paste-format', { conditionalFormatting: false }),
    'conditionalFormatting',
  );
  assert.equal(paste('special-paste-format', { dataValidation: false }), undefined);
  assert.equal(paste('special-paste-besides-border', { borders: false }), 'borders');
  assert.equal(paste('special-paste-besides-border', { dataValidation: false }), 'dataValidation');

  // Optional paste replaces the previous paste by running its stored undo
  // first, so the new mode alone cannot prove a restricted capability absent.
  assert.equal(
    disabledFeatureForCommand(
      { id: 'sheet.command.optional-paste', params: { value: 'PASTE_VALUE' } },
      { dataValidation: false },
    ),
    'dataValidation',
  );
  assert.equal(
    disabledFeatureForCommand(
      { id: 'sheet.command.optional-paste', params: { value: 'PASTE_FORMULA' } },
      { color: false },
    ),
    'color',
  );
});

test('auto-fill preflights every capability its mutation hooks can apply', () => {
  const autoFill = (features: Record<string, boolean>, applyType?: string) =>
    disabledFeatureForCommand(
      {
        id: 'sheet.command.auto-fill',
        params: applyType ? { applyType } : undefined,
      },
      features,
    );

  assert.equal(autoFill({ color: false }, 'COPY'), 'color');
  assert.equal(autoFill({ merge: false }, 'COPY'), 'merge');
  assert.equal(autoFill({ conditionalFormatting: false }, 'SERIES'), 'conditionalFormatting');
  assert.equal(autoFill({ dataValidation: false }, 'ONLY_FORMAT'), 'dataValidation');
  assert.equal(autoFill({ merge: false }), 'merge');

  // NO_FORMAT skips merge/CF/DV hooks. Formatting remains conservatively
  // preflighted because its child payload carries the preserved target style.
  assert.equal(autoFill({ merge: false }, 'NO_FORMAT'), undefined);
  assert.equal(autoFill({ dataValidation: false }, 'NO_FORMAT'), undefined);
  assert.equal(autoFill({ font: false }, 'NO_FORMAT'), 'font');
});

test('structural roots preflight plugin maintenance before changing sheet shape', () => {
  for (const id of [
    'sheet.command.insert-row',
    'sheet.command.insert-row-by-range',
    'sheet.command.insert-col-before',
    'sheet.command.remove-row-by-range',
    'sheet.command.remove-col',
  ]) {
    assert.equal(
      disabledFeatureForCommand({ id }, { tables: false }),
      'tables',
      `expected table preflight: ${id}`,
    );
  }

  for (const id of [
    'sheet.command.move-rows',
    'sheet.command.move-cols',
    'sheet.command.move-range',
  ]) {
    assert.equal(disabledFeatureForCommand({ id }, { filter: false }), 'filter');
    assert.equal(disabledFeatureForCommand({ id }, { tables: false }), undefined);
  }

  for (const id of [
    'sheet.command.insert-range-move-down',
    'sheet.command.delete-range-move-left',
  ]) {
    assert.equal(disabledFeatureForCommand({ id }, { dataValidation: false }), 'dataValidation');
    assert.equal(disabledFeatureForCommand({ id }, { tables: false }), undefined);
    assert.equal(disabledFeatureForCommand({ id }, { filter: false }), undefined);
  }

  assert.equal(
    disabledFeatureForCommand(
      { id: 'sheet.command.insert-row-by-range' },
      { conditionalFormatting: false },
    ),
    'conditionalFormatting',
  );
  assert.equal(
    disabledFeatureForCommand({ id: 'sheet.command.move-range' }, { filter: false }),
    'filter',
  );
  assert.equal(
    disabledFeatureForCommand({ id: 'sheet.command.remove-col' }, { dataValidation: false }),
    'dataValidation',
  );
  assert.equal(
    disabledFeatureForCommand({ id: 'sheet.command.insert-row' }, { merge: false }),
    'merge',
  );

  assert.equal(
    disabledFeatureForCommand({ id: 'sheet.command.reorder-range' }, { dataValidation: false }),
    'dataValidation',
  );
  assert.equal(
    disabledFeatureForCommand(
      { id: 'sheet.command.reorder-range' },
      { conditionalFormatting: false },
    ),
    'conditionalFormatting',
  );
  for (const unrelated of ['merge', 'tables', 'filter']) {
    assert.equal(
      disabledFeatureForCommand({ id: 'sheet.command.reorder-range' }, { [unrelated]: false }),
      undefined,
      `expected reorder to ignore unrelated disabled feature: ${unrelated}`,
    );
  }

  for (const id of ['sheet.command.remove-sheet', 'sheet.command.copy-sheet']) {
    assert.equal(
      disabledFeatureForCommand({ id }, { tables: false }),
      'tables',
      `expected sheet table preflight: ${id}`,
    );
    assert.equal(
      disabledFeatureForCommand({ id }, { conditionalFormatting: false }),
      'conditionalFormatting',
      `expected sheet CF preflight: ${id}`,
    );
    assert.equal(
      disabledFeatureForCommand({ id }, { filter: false }),
      'filter',
      `expected sheet filter preflight: ${id}`,
    );
    assert.equal(
      disabledFeatureForCommand({ id }, { dataValidation: false }),
      'dataValidation',
      `expected sheet DV preflight: ${id}`,
    );
  }
  assert.equal(
    disabledFeatureForCommand({ id: 'sheet.command.copy-sheet' }, { color: false }),
    'color',
  );
  assert.equal(
    disabledFeatureForCommand({ id: 'sheet.command.copy-sheet' }, { merge: false }),
    'merge',
  );
  assert.equal(
    disabledFeatureForCommand({ id: 'sheet.command.remove-sheet' }, { color: false }),
    undefined,
  );
  assert.equal(
    disabledFeatureForCommand({ id: 'sheet.command.remove-sheet' }, { merge: false }),
    undefined,
  );
});

test('refill preflights every capability before undoing the previous fill', () => {
  for (const [feature, expected] of [
    ['font', 'font'],
    ['merge', 'merge'],
    ['conditionalFormatting', 'conditionalFormatting'],
    ['dataValidation', 'dataValidation'],
  ] as const) {
    assert.equal(
      disabledFeatureForCommand({ id: 'sheet.command.refill' }, { [feature]: false }),
      expected,
    );
  }
});

test('undo and redo preflight reactive restrictions before replaying history', () => {
  for (const id of ['univer.command.undo', 'univer.command.redo']) {
    assert.equal(disabledFeatureForCommand({ id }, { history: false }), 'history');
    assert.equal(disabledFeatureForCommand({ id }, { clipboard: false }), 'clipboard');
    assert.equal(disabledFeatureForCommand({ id }, { 'format-painter': false }), 'format-painter');
    assert.equal(disabledFeatureForCommand({ id }, { 'clear-format': false }), 'clear-format');
    assert.equal(disabledFeatureForCommand({ id }, { borders: false }), 'borders');
    assert.equal(disabledFeatureForCommand({ id }, { tables: false }), 'tables');
    assert.equal(
      disabledFeatureForCommand({ id }, { conditionalFormatting: false }),
      'conditionalFormatting',
    );
    assert.equal(disabledFeatureForCommand({ id }, { filter: false }), 'filter');
    assert.equal(disabledFeatureForCommand({ id }, { dataValidation: false }), 'dataValidation');
  }
});

test('serialized trigger metadata cannot bypass the actual command or mutation policy', () => {
  const policy = createCommandPolicy({
    features: () => ({ merge: false }),
    beforeCommand: () => undefined,
  });

  // `executeCommand` and the raw Univer facade both accept arbitrary params.
  // Calling a merge command with an allowed feature's trigger must not relabel
  // the command as paste.
  assert.deepEqual(
    policy({
      id: 'sheet.command.add-worksheet-merge-all',
      params: { trigger: 'univer.command.paste' },
    }),
    { allowed: false, feature: 'merge' },
  );

  // Recorded/replayed mutations are equally caller-controlled. The actual
  // merge mutation remains blocked even when its trigger is forged.
  assert.deepEqual(
    policy({
      id: 'sheet.mutation.add-worksheet-merge',
      params: { trigger: 'univer.command.paste' },
    }),
    { allowed: false, feature: 'merge' },
  );
});

test('serialized trigger metadata cannot hide a disabled style mutation', () => {
  assert.equal(
    disabledFeatureForCommand(
      {
        id: 'sheet.mutation.set-range-values',
        params: {
          trigger: 'univer.command.paste',
          cellValue: { 0: { 0: { s: { bl: 1 } } } },
        },
      },
      { 'font-style': false },
    ),
    'font-style',
  );
});

test('the live policy vetoes non-UI invocation and reacts without remounting', () => {
  let features: Record<string, boolean> | undefined = { merge: false };
  let before: BeforeCommandPolicy | undefined;
  const policy = createCommandPolicy({
    features: () => features,
    beforeCommand: () => before,
  });

  // This is the same command `api.executeCommand`, `api.univer`, a context
  // menu, or a keyboard shortcut reaches at the shared before-command hook.
  assert.deepEqual(policy({ id: 'sheet.command.add-worksheet-merge-all' }), {
    allowed: false,
    feature: 'merge',
  });

  features = { merge: true };
  assert.deepEqual(policy({ id: 'sheet.command.add-worksheet-merge-all' }), { allowed: true });

  before = ({ id }) => id !== 'host.command.block-me';
  assert.deepEqual(policy({ id: 'host.command.block-me' }), { allowed: false });
  assert.deepEqual(policy({ id: 'host.command.allow-me' }), { allowed: true });

  // The callback is id-agnostic: hosts can govern commands introduced by
  // their own plugins without waiting for the SDK's built-in feature map.
  before = ({ id }) => id !== 'host.resource.command.insert-chart';
  assert.deepEqual(policy({ id: 'host.resource.command.insert-chart' }), { allowed: false });
});

test('batch replay runs through the same policy and excludes vetoed steps', async () => {
  const appliedIds: string[] = [];
  const policy = createCommandPolicy({
    features: () => ({ merge: false }),
    beforeCommand: () => undefined,
  });
  const count = await runSteps(
    (id, params) => {
      if (!policy({ id, params }).allowed) return false;
      appliedIds.push(id);
      return true;
    },
    [
      { id: 'sheet.command.set-range-bold' },
      { id: 'sheet.command.add-worksheet-merge-all' },
      { id: 'sheet.command.set-range-italic' },
    ],
  );

  assert.equal(count, 2);
  assert.deepEqual(appliedIds, ['sheet.command.set-range-bold', 'sheet.command.set-range-italic']);
});

test('only authoritative replay and internal formula writes bypass local policy', () => {
  let callbackCalls = 0;
  const policy = createCommandPolicy({
    features: () => ({ merge: false }),
    beforeCommand: () => () => {
      callbackCalls += 1;
      return false;
    },
  });
  const command = { id: 'sheet.mutation.add-worksheet-merge' };

  assert.deepEqual(policy(command, { fromCollab: true }), { allowed: true });
  assert.deepEqual(policy(command, { fromChangeset: true }), { allowed: true });
  assert.deepEqual(policy(command, { onlyLocal: true, fromFormula: true }), { allowed: true });
  assert.deepEqual(
    policy({ id: 'formula.mutation.set-formula-calculation-result' }, { onlyLocal: true }),
    { allowed: true },
  );
  assert.equal(callbackCalls, 0);

  // Neither provenance marker is an authority bypass on its own. The public
  // facade strips both; this pure policy test also protects private engine
  // call sites from accidentally treating one marker as sufficient.
  assert.deepEqual(policy(command, { onlyLocal: true }), {
    allowed: false,
    feature: 'merge',
  });
  assert.deepEqual(policy(command, { fromFormula: true }), {
    allowed: false,
    feature: 'merge',
  });
  assert.equal(callbackCalls, 2);
});

test('host policy observes every local command without overriding disabled features', () => {
  const observed: string[] = [];
  const policy = createCommandPolicy({
    features: () => ({ merge: false }),
    beforeCommand:
      () =>
      ({ id }) => {
        observed.push(id);
        return true;
      },
  });

  assert.deepEqual(policy({ id: 'sheet.command.add-worksheet-merge-all' }), {
    allowed: false,
    feature: 'merge',
  });
  assert.deepEqual(observed, ['sheet.command.add-worksheet-merge-all']);
});

test('host policy receives detached command data and cannot rewrite the engine payload', () => {
  const params = { cellValue: { 0: { 0: { v: 'safe value' } } } };
  const policy = createCommandPolicy({
    features: () => ({ color: false, merge: false }),
    beforeCommand: () => (command) => {
      (command as { id: string }).id = 'host.command.allowed';
      const cell = (
        command.params as {
          cellValue: { 0: { 0: { v: string; s?: { bg: { rgb: string } } } } };
        }
      ).cellValue[0][0];
      cell.s = { bg: { rgb: '#ff0000' } };
    },
  });

  assert.deepEqual(policy({ id: 'sheet.mutation.set-range-values', params }), { allowed: true });
  assert.deepEqual(params, { cellValue: { 0: { 0: { v: 'safe value' } } } });
  assert.deepEqual(policy({ id: 'sheet.command.add-worksheet-merge-all', params }), {
    allowed: false,
    feature: 'merge',
  });
});
