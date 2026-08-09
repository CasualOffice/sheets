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

import { useState } from 'react';
import { CasualSheets, type CasualSheetsAPI } from '@casualoffice/sheets/sheets';
import { startBridge } from '@casualoffice/sheets/collab';
import '@casualoffice/sheets/styles';
import * as Y from 'yjs';
import { emptyWorkbook } from '../snapshot';
import { LOCALES } from '../locale';
import {
  CommandType,
  ICommandService,
  ThemeService,
  type ICommand,
  type IWorkbookData,
  type Univer,
} from '@univerjs/core';
import { UniverSheetsCrosshairHighlightPlugin } from '@univerjs/sheets-crosshair-highlight';

/**
 * Dev-only harness that mounts the SDK's `<CasualSheets>` editor in isolation
 * (no app shell, FileSource, collab, or routing). Reached at `/sdk-harness`.
 *
 * Purpose: give Playwright a way to exercise the SDK editor directly — the app
 * normally renders its own `UniverSheet`, so the published `CasualSheets`
 * component had zero test coverage. This is the verification surface for the
 * SDK-as-full-editor restructure (docs/SDK_AND_DESIGN_PLAN.md), starting with
 * Batch 2's formula engine.
 *
 * Exposes the CasualSheetsAPI ref, a ready flag, and an onChange counter on
 * `window` so specs can drive the editor and observe the snapshot stream
 * deterministically.
 */
export function SdkHarness() {
  const [data] = useState(() => emptyWorkbook());
  const params = new URLSearchParams(window.location.search);
  // `?appearance=dark` mounts the editor in dark mode so the spec can verify it.
  const appearance = params.get('appearance') === 'dark' ? 'dark' : 'light';
  // `?chrome=minimal|full` renders the built-in chrome so the spec can verify it.
  const chromeParam = params.get('chrome');
  const chrome = chromeParam === 'minimal' || chromeParam === 'full' ? chromeParam : 'none';
  // `?disableMerge=1` exercises feature enforcement below the chrome: the
  // Playwright spec dispatches merge through both the stable API and raw
  // FUniver escape hatch, then through executeCommands.
  const features: Record<string, boolean> | undefined =
    params.get('disableMerge') === '1'
      ? { merge: false }
      : params.get('disableTables') === '1'
        ? { tables: false }
        : params.get('disableColor') === '1'
          ? { color: false }
          : undefined;
  const throwOnMerge = params.get('throwOnMerge') === '1';
  const blockFormulaInternal = params.get('blockFormulaInternal') === '1';
  const captureErrors = throwOnMerge || params.get('throwOnObserver') === '1';
  const beforeCommand =
    throwOnMerge || blockFormulaInternal
      ? ({ id }: { id: string }) => {
          if (id === 'sheet.command.add-worksheet-merge-all') {
            throw new Error('host policy exploded');
          }
          if (blockFormulaInternal && id === 'formula.mutation.set-formula-calculation-result') {
            return false;
          }
        }
      : undefined;
  // `?beforeCreate=crosshair` exercises the onBeforeCreateUnit escape hatch by
  // registering a plugin the SDK doesn't bundle (crosshair-highlight) — the
  // spec then asserts its command registered, proving the hook works.
  const beforeCreate =
    params.get('beforeCreate') === 'crosshair'
      ? (univer: Univer) => {
          univer.registerPlugin(UniverSheetsCrosshairHighlightPlugin);
        }
      : undefined;
  // `?formulaWorker=1` exercises off-main formula compute via the standard
  // Univer formula worker — the spec asserts compute still works through it.
  const formula =
    params.get('formulaWorker') === '1'
      ? {
          worker: new Worker(new URL('../univer/formula-worker.ts', import.meta.url), {
            type: 'module',
            name: 'formula-worker',
          }),
        }
      : undefined;
  return (
    <div data-testid="sdk-harness" style={{ position: 'fixed', inset: 0 }}>
      <CasualSheets
        initialData={data}
        locales={LOCALES}
        appearance={appearance}
        chrome={chrome}
        features={features}
        onBeforeCommand={beforeCommand}
        onBeforeCreateUnit={beforeCreate}
        formula={formula}
        onError={
          captureErrors
            ? (error) => {
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                const w = window as any;
                w.__sdkHarnessErrors = [...(w.__sdkHarnessErrors ?? []), error.message];
              }
            : undefined
        }
        onReady={(api: CasualSheetsAPI) => {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          (window as any).__sdkHarnessAPI = api;
          // Expose hasCommand so specs can check lazy plugins registered without
          // importing redi tokens into page context.
          const injector = (api.univer as unknown as { _injector?: { get(t: unknown): unknown } })
            ._injector;
          const svc = injector?.get(ICommandService) as
            | {
                hasCommand(id: string): boolean;
                registerCommand(command: ICommand): { dispose(): void };
                syncExecuteCommand(id: string, params?: object): boolean;
              }
            | undefined;
          if (throwOnMerge && svc && !svc.hasCommand('test.command.nested-merge')) {
            svc.registerCommand({
              id: 'test.command.nested-merge',
              type: CommandType.COMMAND,
              handler: () => svc.syncExecuteCommand('sheet.command.add-worksheet-merge-all'),
            });
          }
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          (window as any).__sdkHarnessHasCommand = (id: string) => svc?.hasCommand(id) ?? false;
          // Test-only driver for a real Yjs → startBridge replay. Public
          // api.univer dispatch strips replay provenance; the bridge's private
          // command-service path intentionally retains it for convergence.
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          (window as any).__sdkHarnessReplayCommand = async (
            id: string,
            commandParams?: object,
          ) => {
            const receiver = new Y.Doc();
            const sender = new Y.Doc();
            const bridge = startBridge(api.univer, receiver, { role: 'view' });
            sender
              .getArray('ops')
              .push([
                {
                  v: 2,
                  kind: 'op',
                  c: String(sender.clientID),
                  s: 0,
                  t: Date.now(),
                  id,
                  p: commandParams,
                },
              ]);
            Y.applyUpdate(receiver, Y.encodeStateAsUpdate(sender));
            await bridge.whenReplaySettled();
            const failures = bridge.getReplayFailures();
            bridge.dispose();
            sender.destroy();
            receiver.destroy();
            return { failures };
          };
          // Expose Univer's dark-mode flag so the appearance spec can assert it.
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          (window as any).__sdkHarnessIsDark = () => {
            const injector = (api.univer as unknown as { _injector?: { get(t: unknown): unknown } })
              ._injector;
            const svc = injector?.get(ThemeService) as { darkMode: boolean } | undefined;
            return svc?.darkMode ?? false;
          };
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          (window as any).__sdkHarnessReady = true;
        }}
        onChange={(snapshot: IWorkbookData) => {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const w = window as any;
          w.__sdkHarnessChangeCount = (w.__sdkHarnessChangeCount ?? 0) + 1;
          w.__sdkHarnessLastSnapshot = snapshot;
        }}
        onSave={(snapshot: IWorkbookData) => {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const w = window as any;
          w.__sdkHarnessSaveCount = (w.__sdkHarnessSaveCount ?? 0) + 1;
          w.__sdkHarnessLastSaved = snapshot;
        }}
        onExit={(snapshot: IWorkbookData) => {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          (window as any).__sdkHarnessExited = !!snapshot;
        }}
      />
    </div>
  );
}
