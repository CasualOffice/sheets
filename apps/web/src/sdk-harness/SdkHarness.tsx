import { useState } from 'react';
import { CasualSheets, type CasualSheetsAPI } from '@casualoffice/sheets/sheets';
import '@casualoffice/sheets/styles';
import { emptyWorkbook } from '../snapshot';
import { LOCALES } from '../locale';
import type { IWorkbookData } from '@univerjs/core';

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
  return (
    <div data-testid="sdk-harness" style={{ position: 'fixed', inset: 0 }}>
      <CasualSheets
        initialData={data}
        locales={LOCALES}
        onReady={(api: CasualSheetsAPI) => {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          (window as any).__sdkHarnessAPI = api;
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          (window as any).__sdkHarnessReady = true;
        }}
        onChange={(snapshot: IWorkbookData) => {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const w = window as any;
          w.__sdkHarnessChangeCount = (w.__sdkHarnessChangeCount ?? 0) + 1;
          w.__sdkHarnessLastSnapshot = snapshot;
        }}
      />
    </div>
  );
}
