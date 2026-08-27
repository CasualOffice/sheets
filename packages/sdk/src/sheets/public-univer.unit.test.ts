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

import { createPublicUniverFacade, sanitizePublicExecutionOptions } from './public-univer';

test('public execution options cannot forge authoritative replay provenance', () => {
  assert.deepEqual(
    sanitizePublicExecutionOptions({
      fromCollab: true,
      fromChangeset: true,
      fromFormula: true,
      applyFormulaCalculationResult: true,
      onlyLocal: true,
      syncOnly: true,
    }),
    {},
  );
});

test('public facade guards async and sync commands without fluent escape hatches', async () => {
  const received: unknown[] = [];
  const facade = {
    marker: 42,
    readMarker() {
      return this.marker;
    },
    executeCommand(_id: string, _params?: object, options?: object) {
      received.push(options);
      return true;
    },
    syncExecuteCommand(_id: string, _params?: object, options?: object) {
      received.push(options);
      return true;
    },
    showMessage() {
      return this;
    },
    async asyncFluent() {
      return this;
    },
  };
  const publicFacade = createPublicUniverFacade(facade);

  assert.equal(publicFacade.readMarker(), 42);
  assert.equal(publicFacade.readMarker, publicFacade.readMarker, 'bound method identity is stable');
  assert.equal(
    await publicFacade.executeCommand(
      'local.command',
      {},
      {
        fromCollab: true,
        syncOnly: true,
      },
    ),
    true,
  );
  assert.equal(
    publicFacade.syncExecuteCommand(
      'local.sync-command',
      {},
      { fromChangeset: true, syncOnly: true },
    ),
    true,
  );
  assert.deepEqual(received, [{}, {}]);
  assert.equal(publicFacade.showMessage(), publicFacade);
  assert.equal(await publicFacade.asyncFluent(), publicFacade);
  assert.equal(
    publicFacade
      .showMessage()
      .syncExecuteCommand('local.chained-command', {}, { fromCollab: true, chained: true }),
    true,
  );
  assert.deepEqual(received[2], { chained: true });
});

test('public command listeners receive provenance-isolated option clones', () => {
  let beforeListener: ((command: object, options?: object) => void) | undefined;
  let afterListener: ((command: object, options?: object) => void) | undefined;
  const eventListeners = new Map<string, (event: Record<string, unknown>) => void>();
  const facade = {
    onBeforeCommandExecute(callback: (command: object, options?: object) => void) {
      beforeListener = callback;
      return { dispose() {} };
    },
    onCommandExecuted(callback: (command: object, options?: object) => void) {
      afterListener = callback;
      return { dispose() {} };
    },
    addEvent(event: string, callback: (value: Record<string, unknown>) => void) {
      eventListeners.set(event, callback);
      return { dispose() {} };
    },
  };
  const publicFacade = createPublicUniverFacade(facade);

  publicFacade.onBeforeCommandExecute((command, options) => {
    const params = command as { params?: { cellValue?: { value?: string } } };
    if (params.params?.cellValue) params.params.cellValue.value = 'rewritten';
    Object.assign(options ?? {}, { fromCollab: true, onlyLocal: true });
  });
  publicFacade.onCommandExecuted((_command, options) => {
    Object.assign(options ?? {}, { fromChangeset: true, onlyLocal: true });
  });
  publicFacade.addEvent('BeforeCommandExecute', (event) => {
    const params = event.params as { cellValue?: { value?: string } } | undefined;
    if (params?.cellValue) params.cellValue.value = 'event rewrite';
    Object.assign((event.options as object | undefined) ?? {}, { fromFormula: true });
    event.cancel = true;
  });
  publicFacade.addEvent('CommandExecuted', (event) => {
    Object.assign((event.options as object | undefined) ?? {}, { onlyLocal: true });
  });

  const beforeOptions = { custom: true };
  const beforeCommand = { id: 'before', params: { cellValue: { value: 'original' } } };
  beforeListener?.(beforeCommand, beforeOptions);
  assert.deepEqual(beforeOptions, { custom: true });
  assert.equal(beforeCommand.params.cellValue.value, 'original');

  const afterOptions = { custom: true };
  afterListener?.({ id: 'after' }, afterOptions);
  assert.deepEqual(afterOptions, { custom: true });

  const beforeEvent = {
    params: { cellValue: { value: 'original' } },
    options: { custom: true },
    cancel: false,
  };
  eventListeners.get('BeforeCommandExecute')?.(beforeEvent);
  assert.deepEqual(beforeEvent.options, { custom: true });
  assert.equal(beforeEvent.params.cellValue.value, 'original');
  assert.equal(beforeEvent.cancel, true, 'documented event cancellation still reaches Univer');

  const afterEvent = { options: { custom: true } };
  eventListeners.get('CommandExecuted')?.(afterEvent);
  assert.deepEqual(afterEvent.options, { custom: true });
});
