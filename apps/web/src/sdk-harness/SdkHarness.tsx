import { useState } from 'react';
import { CasualSheets } from '@casualoffice/sheets/sheets';
import '@casualoffice/sheets/styles';
import { emptyWorkbook } from '../snapshot';
import type { FUniver } from '@univerjs/core/facade';

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
 * Exposes the FUniver facade + a ready flag on `window` so specs can drive the
 * editor (enter a formula, read a computed cell) deterministically.
 */
export function SdkHarness() {
  const [data] = useState(() => emptyWorkbook());
  return (
    <div data-testid="sdk-harness" style={{ position: 'fixed', inset: 0 }}>
      <CasualSheets
        initialData={data}
        onReady={(api: FUniver) => {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          (window as any).__sdkHarnessAPI = api;
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          (window as any).__sdkHarnessReady = true;
        }}
      />
    </div>
  );
}
