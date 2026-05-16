import { useMemo, useState, type ReactNode } from 'react';
import type { FUniver } from '@univerjs/core/facade';
import { UniverContext, type UniverCtxValue } from './univer-context';

/**
 * Top-level wrapper that owns the FUniver context state.
 */
export function UniverRoot({ children }: { children: ReactNode }) {
  const [api, setApi] = useState<FUniver | null>(null);
  const value = useMemo<UniverCtxValue>(() => ({ api, setApi }), [api]);
  return <UniverContext.Provider value={value}>{children}</UniverContext.Provider>;
}
