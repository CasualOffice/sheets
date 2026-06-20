import { useState } from 'react';
import { CasualSheets, type CasualSheetsAPI } from '@casualoffice/sheets/sheets';
import '@casualoffice/sheets/styles';
import { emptyWorkbook } from '../snapshot';
import { LOCALES } from '../locale';
import { ICommandService, ThemeService, type IWorkbookData } from '@univerjs/core';

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
  return (
    <div data-testid="sdk-harness" style={{ position: 'fixed', inset: 0 }}>
      <CasualSheets
        initialData={data}
        locales={LOCALES}
        appearance={appearance}
        chrome={chrome}
        onReady={(api: CasualSheetsAPI) => {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          (window as any).__sdkHarnessAPI = api;
          // Expose hasCommand so specs can check lazy plugins registered without
          // importing redi tokens into page context.
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          (window as any).__sdkHarnessHasCommand = (id: string) => {
            const injector = (api.univer as unknown as { _injector?: { get(t: unknown): unknown } })
              ._injector;
            const svc = injector?.get(ICommandService) as
              | { hasCommand(id: string): boolean }
              | undefined;
            return svc?.hasCommand(id) ?? false;
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
