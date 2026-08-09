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
import * as Y from 'yjs';
import type { ICommandInfo, IExecutionOptions, IWorkbookData } from '@univerjs/core';
import type { CasualSheetsAPI } from '../sheets/api';
import { createPublicUniverFacade } from '../sheets/public-univer';
import { startBridge, SYNCED_MUTATIONS, type BridgeHandle } from './bridge';
import {
  COLLAB_LOG_PROTOCOL_VERSION,
  type MutationRecord,
  type SnapshotRecord,
} from './replay-plan';

const SET_VALUES = 'sheet.mutation.set-range-values';

type CommandListener = (info: ICommandInfo, options?: IExecutionOptions) => void;
type Cells = Record<string, unknown>;

type Harness = {
  api: CasualSheetsAPI;
  doc: Y.Doc;
  bridge: BridgeHandle;
  cells: () => Cells;
  calls: string[];
  emitLocal: (row: number, col: number, value: unknown) => void;
  emitLocalWithForgedReplayFlag: (row: number, col: number, value: unknown) => void;
  emitLocalWithOptions: (
    row: number,
    col: number,
    value: unknown,
    options: IExecutionOptions,
  ) => void;
  emitMutation: (id: string, params?: unknown, options?: IExecutionOptions) => void;
  emitSplitLocal: (row: number, col: number, value: unknown) => void;
  emitUnsupported: (id: string) => void;
  emitUnsupportedWithOptions: (id: string, options: IExecutionOptions) => void;
  applyParams: (params: unknown) => void;
  setExecute: (fn: (id: string, params: unknown) => Promise<unknown>) => void;
  setSnapshotApply: (fn: (wb: IWorkbookData, commit: () => void) => void | Promise<void>) => void;
  setContentAvailable: (available: boolean) => void;
};

function createHarness(resources?: IWorkbookData['resources']): Harness {
  const doc = new Y.Doc();
  const mutationListeners = new Set<CommandListener>();
  const beforeListeners = new Set<CommandListener>();
  const calls: string[] = [];
  let cells: Cells = {};
  let contentAvailable = true;
  let snapshotApply: (wb: IWorkbookData, commit: () => void) => void | Promise<void> = (
    _wb,
    commit,
  ) => commit();
  let executeImpl = async (id: string, params: unknown): Promise<unknown> => {
    calls.push(id);
    applyCellParams(cells, params);
    return true;
  };

  const commandService = {
    onMutationExecutedForCollab(listener: CommandListener) {
      mutationListeners.add(listener);
      return { dispose: () => mutationListeners.delete(listener) };
    },
    beforeCommandExecuted(listener: CommandListener) {
      beforeListeners.add(listener);
      return { dispose: () => beforeListeners.delete(listener) };
    },
    async executeCommand(id: string, params: unknown, options?: IExecutionOptions) {
      const result = await executeImpl(id, params);
      for (const listener of mutationListeners) {
        listener({ id, params } as ICommandInfo, options);
      }
      return result;
    },
  };
  const facade = {
    _injector: { get: () => commandService },
    getActiveWorkbook: () => ({
      getId: () => 'local-wb',
      // Deliberately lossy: compaction must never call this raw save path.
      save: () => workbookFromCells(cells),
      getActiveSheet: () => ({ getSheetId: () => 'sheet-1' }),
      getSheets: () => [],
    }),
  };
  const api = {
    univer: facade,
    getContent: () =>
      contentAvailable ? workbookFromCells(cells, resources) : (null as IWorkbookData | null),
    setContent: (wb: IWorkbookData) =>
      snapshotApply(wb, () => {
        cells = cellsFromWorkbook(wb);
      }),
  } as unknown as CasualSheetsAPI;

  const bridge = startBridge(api, doc, {
    awareness: { getStates: () => new Map([[doc.clientID, {}]]) },
    compaction: 'manual',
  });

  return {
    api,
    doc,
    bridge,
    cells: () => ({ ...cells }),
    calls,
    emitLocal: (row, col, value) => {
      const params = cellParams(row, col, value);
      applyCellParams(cells, params);
      for (const listener of mutationListeners)
        listener({ id: SET_VALUES, params } as ICommandInfo);
    },
    emitLocalWithForgedReplayFlag: (row, col, value) => {
      const params = cellParams(row, col, value);
      applyCellParams(cells, params);
      for (const listener of mutationListeners) {
        listener({ id: SET_VALUES, params } as ICommandInfo, { fromCollab: true });
      }
    },
    emitLocalWithOptions: (row, col, value, options) => {
      const params = cellParams(row, col, value);
      applyCellParams(cells, params);
      for (const listener of mutationListeners) {
        listener({ id: SET_VALUES, params } as ICommandInfo, options);
      }
    },
    emitMutation: (id, params = {}, options) => {
      for (const listener of mutationListeners) {
        listener({ id, params } as ICommandInfo, options);
      }
    },
    emitSplitLocal: (row, col, value) => {
      const params = { ...(cellParams(row, col, value) as object), __splitChunk__: true };
      applyCellParams(cells, params);
      for (const listener of mutationListeners)
        listener({ id: SET_VALUES, params } as ICommandInfo);
    },
    emitUnsupported: (id) => {
      for (const listener of mutationListeners) listener({ id, params: {} } as ICommandInfo);
    },
    emitUnsupportedWithOptions: (id, options) => {
      for (const listener of mutationListeners) {
        listener({ id, params: {} } as ICommandInfo, options);
      }
    },
    applyParams: (params) => applyCellParams(cells, params),
    setExecute: (fn) => {
      executeImpl = async (id, params) => {
        calls.push(id);
        return fn(id, params);
      };
    },
    setSnapshotApply: (fn) => {
      snapshotApply = fn;
    },
    setContentAvailable: (available) => {
      contentAvailable = available;
    },
  };
}

test('two offline clients keep both concurrent different-cell edits', async () => {
  const a = createHarness();
  const b = createHarness();
  await Promise.all([a.bridge.whenReplaySettled(), b.bridge.whenReplaySettled()]);

  a.emitLocal(0, 0, 'from-a');
  b.emitLocal(0, 1, 'from-b');
  await nextMicrotask();
  await Promise.all([a.bridge.whenReplaySettled(), b.bridge.whenReplaySettled()]);

  const updateA = Y.encodeStateAsUpdate(a.doc);
  const updateB = Y.encodeStateAsUpdate(b.doc);
  Y.applyUpdate(a.doc, updateB);
  Y.applyUpdate(b.doc, updateA);
  await Promise.all([a.bridge.whenReplaySettled(), b.bridge.whenReplaySettled()]);

  assert.deepEqual(a.cells(), { '0:0': 'from-a', '0:1': 'from-b' });
  assert.deepEqual(b.cells(), a.cells());
  a.bridge.dispose();
  b.bridge.dispose();
});

test('concurrent same-cell writes converge by restoring and replaying canonical order', async () => {
  const a = createHarness();
  const b = createHarness();
  await Promise.all([a.bridge.whenReplaySettled(), b.bridge.whenReplaySettled()]);

  a.emitLocal(0, 0, 'a');
  b.emitLocal(0, 0, 'b');
  await nextMicrotask();
  await Promise.all([a.bridge.whenReplaySettled(), b.bridge.whenReplaySettled()]);
  const updateA = Y.encodeStateAsUpdate(a.doc);
  const updateB = Y.encodeStateAsUpdate(b.doc);
  Y.applyUpdate(a.doc, updateB);
  Y.applyUpdate(b.doc, updateA);
  await Promise.all([a.bridge.whenReplaySettled(), b.bridge.whenReplaySettled()]);

  assert.deepEqual(a.cells(), b.cells());
  assert.ok(a.cells()['0:0'] === 'a' || a.cells()['0:0'] === 'b');
  a.bridge.dispose();
  b.bridge.dispose();
});

test('a local edit made during an async rebuild is replayed after the swap', async () => {
  const h = createHarness();
  await h.bridge.whenReplaySettled();
  h.emitLocal(0, 1, 'existing-local');
  await h.bridge.whenReplaySettled();

  let signalStarted!: () => void;
  const started = new Promise<void>((resolve) => {
    signalStarted = resolve;
  });
  let releaseSwap!: () => void;
  const swapReleased = new Promise<void>((resolve) => {
    releaseSwap = resolve;
  });
  let firstSwap = true;
  h.setSnapshotApply(async (_wb, commit) => {
    if (firstSwap) {
      firstSwap = false;
      signalStarted();
      await swapReleased;
    }
    commit();
  });

  // Explicit insertion before the already-applied local record forces a reset.
  h.doc.getArray<MutationRecord>('ops').insert(0, [remoteOp('peer', 0, 0, 0, 'remote')]);
  await started;
  h.emitLocal(0, 2, 'during-swap');
  releaseSwap();
  await h.bridge.whenReplaySettled();

  assert.deepEqual(h.cells(), {
    '0:0': 'remote',
    '0:1': 'existing-local',
    '0:2': 'during-swap',
  });
  h.bridge.dispose();
});

test('replay awaits each async command before starting the next', async () => {
  const h = createHarness();
  await h.bridge.whenReplaySettled();
  const order: string[] = [];
  let releaseFirst!: () => void;
  const firstDone = new Promise<void>((resolve) => {
    releaseFirst = resolve;
  });
  h.setExecute(async (_id, params) => {
    const value = readOnlyCellValue(params);
    order.push(`start:${String(value)}`);
    if (value === 'first') await firstDone;
    h.applyParams(params);
    order.push(`end:${String(value)}`);
    return true;
  });

  h.doc
    .getArray<MutationRecord>('ops')
    .push([remoteOp('peer', 0, 0, 0, 'first'), remoteOp('peer', 1, 0, 1, 'second')]);
  await nextMicrotask();
  assert.deepEqual(order, ['start:first']);
  releaseFirst();
  await h.bridge.whenReplaySettled();
  assert.deepEqual(order, ['start:first', 'end:first', 'start:second', 'end:second']);
  h.bridge.dispose();
});

test('split chunks retain separate Yjs changeset boundaries', async () => {
  const h = createHarness();
  await h.bridge.whenReplaySettled();
  const log = h.doc.getArray<MutationRecord>('ops');
  let transactions = 0;
  const observe = () => {
    transactions += 1;
  };
  log.observe(observe);
  h.emitSplitLocal(0, 0, 'chunk-1');
  h.emitSplitLocal(1, 0, 'chunk-2');
  await h.bridge.whenReplaySettled();
  assert.equal(log.length, 2);
  assert.equal(transactions, 2);
  log.unobserve(observe);
  h.bridge.dispose();
});

test('a public caller cannot forge fromCollab to hide a local mutation', async () => {
  const h = createHarness();
  await h.bridge.whenReplaySettled();
  h.emitLocalWithForgedReplayFlag(0, 0, 'must-be-recorded');
  await h.bridge.whenReplaySettled();
  assert.equal(h.doc.getArray<MutationRecord>('ops').length, 1);
  assert.equal(h.bridge.isReplayBlocked(), false);
  assert.deepEqual(h.cells(), { '0:0': 'must-be-recorded' });
  h.bridge.dispose();
});

test('public observers cannot capture and reuse bridge replay provenance', async () => {
  let engineListener: ((command: object, options?: object) => void) | undefined;
  let dispatchedOptions: object | undefined;
  const rawFacade = {
    onCommandExecuted(listener: (command: object, options?: object) => void) {
      engineListener = listener;
      return { dispose() {} };
    },
    async executeCommand(_id: string, _params?: object, options?: object) {
      dispatchedOptions = options;
      return true;
    },
  };
  const publicFacade = createPublicUniverFacade(rawFacade);
  let capturedOptions: object | undefined;
  publicFacade.onCommandExecuted((_command, options) => {
    capturedOptions = options;
  });

  const replayToken = { fromCollab: true };
  engineListener?.({ id: SET_VALUES }, replayToken);
  assert.notEqual(capturedOptions, replayToken);
  assert.deepEqual(capturedOptions, {});

  await publicFacade.executeCommand('local.command', {}, capturedOptions);
  assert.notEqual(dispatchedOptions, replayToken);
  assert.deepEqual(dispatchedOptions, {});
});

test('non-replicating runtime and snapshot mutations never enter the collaboration log', async () => {
  const h = createHarness();
  await h.bridge.whenReplaySettled();

  h.emitLocalWithOptions(0, 0, 'formula-cache', {
    onlyLocal: true,
    fromFormula: true,
    applyFormulaCalculationResult: true,
  });
  h.emitLocalWithOptions(0, 1, 'snapshot-value', { fromChangeset: true });
  h.emitUnsupportedWithOptions('formula.mutation.set-formula-calculation-result', {
    onlyLocal: true,
  });
  await h.bridge.whenReplaySettled();

  assert.equal(h.doc.getArray<MutationRecord>('ops').length, 0);
  assert.equal(h.bridge.isReplayBlocked(), false);
  assert.deepEqual(h.cells(), { '0:0': 'formula-cache', '0:1': 'snapshot-value' });
  h.bridge.dispose();
});

test('the allowlist uses the mutation ids emitted by Univer 0.25', () => {
  for (const id of [
    'sheet.mutation.remove-rows',
    'sheet.mutation.move-columns',
    'sheet.mutation.set.numfmt',
    'sheet.mutation.remove.numfmt',
    'sheet.mutation.set-worksheet-row-auto-height',
    'sheets.mutation.add-hyper-link',
    'sheets.mutation.remove-hyper-link',
    'sheets.mutation.update-hyper-link',
  ]) {
    assert.equal(SYNCED_MUTATIONS.has(id), true, `${id} should be synchronized`);
  }

  for (const staleId of [
    'sheet.mutation.remove-row',
    'sheet.mutation.move-cols',
    'sheet.mutation.add-hyper-link',
  ]) {
    assert.equal(SYNCED_MUTATIONS.has(staleId), false, `${staleId} is not a Univer 0.25 id`);
  }
});

test('current structural, formatting, and hyperlink mutations replay on a peer', async () => {
  const a = createHarness();
  const b = createHarness();
  await Promise.all([a.bridge.whenReplaySettled(), b.bridge.whenReplaySettled()]);
  b.setExecute(async () => true);
  const ids = [
    'sheet.mutation.remove-rows',
    'sheet.mutation.move-columns',
    'sheet.mutation.set.numfmt',
    'sheet.mutation.remove.numfmt',
    'sheet.mutation.set-worksheet-row-auto-height',
    'sheets.mutation.add-hyper-link',
    'sheets.mutation.remove-hyper-link',
    'sheets.mutation.update-hyper-link',
  ];

  for (const id of ids) {
    a.emitMutation(id, { unitId: 'sender-wb', subUnitId: 'sheet-1' });
  }
  await a.bridge.whenReplaySettled();
  Y.applyUpdate(b.doc, Y.encodeStateAsUpdate(a.doc));
  await b.bridge.whenReplaySettled();

  assert.deepEqual(b.calls, ids);
  assert.equal(b.bridge.isReplayBlocked(), false);
  a.bridge.dispose();
  b.bridge.dispose();
});

test('unknown and failed records block the head without advancing later mutations', async () => {
  const h = createHarness();
  await h.bridge.whenReplaySettled();
  h.doc
    .getArray<MutationRecord>('ops')
    .push([
      { ...remoteOp('peer', 0, 0, 0, 'unknown'), id: 'sheet.mutation.future-command' },
      remoteOp('peer', 1, 0, 1, 'must-not-run'),
    ]);
  await h.bridge.whenReplaySettled();
  assert.equal(h.bridge.isReplayBlocked(), true);
  assert.equal(h.bridge.getReplayFailures(), 1);
  assert.deepEqual(h.cells(), {});
  assert.deepEqual(h.calls, []);
  h.bridge.dispose();

  const retry = createHarness();
  await retry.bridge.whenReplaySettled();
  let fail = true;
  const attemptedValues: unknown[] = [];
  retry.setExecute(async (_id, params) => {
    attemptedValues.push(readOnlyCellValue(params));
    if (fail) return false;
    retry.applyParams(params);
    return true;
  });
  retry.doc
    .getArray<MutationRecord>('ops')
    .push([remoteOp('peer', 0, 0, 0, 'recovered'), remoteOp('peer', 1, 0, 1, 'after-recovery')]);
  await retry.bridge.whenReplaySettled();
  assert.equal(retry.bridge.isReplayBlocked(), true);
  assert.ok(attemptedValues.length >= 1);
  assert.ok(
    attemptedValues.every((value) => value === 'recovered'),
    'the later mutation must stay behind the failed head',
  );
  fail = false;
  await retry.bridge.whenReplaySettled();
  assert.equal(retry.bridge.isReplayBlocked(), false);
  assert.deepEqual(attemptedValues.slice(-2), ['recovered', 'after-recovery']);
  assert.deepEqual(retry.cells(), { '0:0': 'recovered', '0:1': 'after-recovery' });
  retry.bridge.dispose();
});

test('forced compaction preserves opaque resources and refuses null or divergent state', async () => {
  const opaque = [{ name: 'HOST_OPAQUE_RESOURCE', data: 'keep-me' }];
  const h = createHarness(opaque);
  await h.bridge.whenReplaySettled();
  for (let row = 0; row < 200; row += 1) h.emitLocal(row, 0, row);
  // No microtask yield: whenReplaySettled must drain the locally queued batch,
  // not attach to an already-resolved startup flight and return early.
  await h.bridge.whenReplaySettled();
  assert.equal(h.bridge.forceCompact(), true);
  const compacted = h.doc.getArray<SnapshotRecord>('ops').get(0);
  assert.equal(compacted.kind, 'snapshot');
  assert.deepEqual(compacted.wb.resources, opaque);
  h.bridge.dispose();

  const noContent = createHarness(opaque);
  await noContent.bridge.whenReplaySettled();
  for (let row = 0; row < 200; row += 1) noContent.emitLocal(row, 0, row);
  await nextMicrotask();
  await noContent.bridge.whenReplaySettled();
  noContent.setContentAvailable(false);
  assert.equal(noContent.bridge.forceCompact(), false);
  assert.equal(noContent.doc.getArray('ops').length, 200);
  noContent.bridge.dispose();

  const blocked = createHarness(opaque);
  await blocked.bridge.whenReplaySettled();
  const records: MutationRecord[] = [
    { ...remoteOp('peer', 0, 0, 0, 'unknown'), id: 'sheet.mutation.future-command' },
  ];
  for (let row = 1; row < 200; row += 1) {
    records.push(remoteOp('peer', row, row, 0, row));
  }
  blocked.doc.getArray<MutationRecord>('ops').push(records);
  await blocked.bridge.whenReplaySettled();
  assert.equal(blocked.bridge.forceCompact(), false);
  assert.equal(blocked.doc.getArray('ops').length, 200);
  blocked.bridge.dispose();

  const locallyDivergent = createHarness(opaque);
  await locallyDivergent.bridge.whenReplaySettled();
  for (let row = 0; row < 200; row += 1) locallyDivergent.emitLocal(row, 0, row);
  await nextMicrotask();
  await locallyDivergent.bridge.whenReplaySettled();
  locallyDivergent.emitUnsupported('sheet.mutation.unversioned-local-state');
  assert.equal(locallyDivergent.bridge.isReplayBlocked(), true);
  assert.equal(locallyDivergent.bridge.getReplayFailures(), 0);
  assert.equal(locallyDivergent.bridge.forceCompact(), false);
  locallyDivergent.bridge.dispose();
});

test('compaction defaults to off and explicit off still replays collaboration', async () => {
  for (const compaction of [undefined, 'off'] as const) {
    const doc = new Y.Doc();
    const calls: string[] = [];
    doc
      .getArray<MutationRecord>('ops')
      .push(Array.from({ length: 200 }, (_, row) => remoteOp('peer', row, row, 0, row)));
    const facade = {
      _injector: {
        get: () => ({
          onMutationExecutedForCollab: () => ({ dispose() {} }),
          beforeCommandExecuted: () => ({ dispose() {} }),
          executeCommand: async (id: string) => {
            calls.push(id);
            return true;
          },
        }),
      },
      getActiveWorkbook: () => ({ getId: () => 'wb', getActiveSheet: () => null }),
    };
    const api = {
      univer: facade,
      getContent: () => workbookFromCells({}),
      setContent: () => undefined,
    } as unknown as CasualSheetsAPI;
    const bridge = startBridge(api, doc, {
      awareness: { getStates: () => new Map([[doc.clientID, {}]]) },
      ...(compaction ? { compaction } : {}),
    });
    await bridge.whenReplaySettled();
    assert.equal(calls.length, 200);
    assert.equal(bridge.forceCompact(), false);
    assert.equal(doc.getArray('ops').length, 200);
    bridge.dispose();
  }
});

function remoteOp(
  client: string,
  sequence: number,
  row: number,
  col: number,
  value: unknown,
): MutationRecord {
  return {
    v: COLLAB_LOG_PROTOCOL_VERSION,
    kind: 'op',
    c: client,
    s: sequence,
    t: sequence,
    id: SET_VALUES,
    p: cellParams(row, col, value),
  };
}

function cellParams(row: number, col: number, value: unknown): unknown {
  return {
    unitId: 'sender-wb',
    subUnitId: 'sheet-1',
    cellValue: { [row]: { [col]: { v: value } } },
  };
}

function applyCellParams(target: Cells, params: unknown): void {
  const valueMap = (params as { cellValue?: Record<string, Record<string, { v?: unknown }>> })
    .cellValue;
  if (!valueMap) throw new Error('missing cellValue');
  for (const [row, columns] of Object.entries(valueMap)) {
    for (const [col, cell] of Object.entries(columns)) target[`${row}:${col}`] = cell.v;
  }
}

function readOnlyCellValue(params: unknown): unknown {
  const rows = Object.values(
    (params as { cellValue: Record<string, Record<string, { v?: unknown }>> }).cellValue,
  );
  const cell = Object.values(rows[0] ?? {})[0];
  return cell?.v;
}

function workbookFromCells(cells: Cells, resources?: IWorkbookData['resources']): IWorkbookData {
  const cellData: Record<number, Record<number, { v: unknown }>> = {};
  for (const [key, value] of Object.entries(cells)) {
    const [row, col] = key.split(':').map(Number);
    cellData[row] ??= {};
    cellData[row][col] = { v: value };
  }
  return {
    id: 'local-wb',
    name: 'Workbook',
    sheetOrder: ['sheet-1'],
    sheets: {
      'sheet-1': {
        id: 'sheet-1',
        name: 'Sheet1',
        rowCount: 1000,
        columnCount: 100,
        cellData,
      },
    },
    resources,
  } as IWorkbookData;
}

function cellsFromWorkbook(wb: IWorkbookData): Cells {
  const result: Cells = {};
  const data = wb.sheets?.['sheet-1']?.cellData ?? {};
  for (const [row, columns] of Object.entries(data)) {
    for (const [col, cell] of Object.entries(columns ?? {})) {
      result[`${row}:${col}`] = (cell as { v?: unknown }).v;
    }
  }
  return result;
}

async function nextMicrotask(): Promise<void> {
  await Promise.resolve();
  await Promise.resolve();
}
