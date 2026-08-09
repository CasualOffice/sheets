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

import { useCallback, useMemo, useState, type ReactNode } from 'react';
import type { FUniver } from '@univerjs/core/facade';
import type { CasualSheetsAPI } from '@casualoffice/sheets/sheets';
import { UniverContext, type UniverCtxValue } from './univer-context';

/**
 * Top-level wrapper that owns the FUniver context state.
 */
export function UniverRoot({ children }: { children: ReactNode }) {
  const [api, setApi] = useState<FUniver | null>(null);
  const [sheetsApi, setSheetsApi] = useState<CasualSheetsAPI | null>(null);
  const publishApi = useCallback(
    (nextApi: FUniver | null, nextSheetsApi: CasualSheetsAPI | null = null) => {
      setApi(nextApi);
      setSheetsApi(nextSheetsApi);
    },
    [],
  );
  const value = useMemo<UniverCtxValue>(
    () => ({ api, sheetsApi, setApi: publishApi }),
    [api, sheetsApi, publishApi],
  );
  return <UniverContext.Provider value={value}>{children}</UniverContext.Provider>;
}
