import { useContext } from 'react';
import type { FUniver } from '@univerjs/core/facade';
import { UniverContext } from './univer-context';

/** Returns the Facade API once ready, or `null` while Univer is mounting. */
export function useUniverAPI(): FUniver | null {
  const ctx = useContext(UniverContext);
  if (!ctx) {
    throw new Error('useUniverAPI must be used inside <UniverProvider>');
  }
  return ctx.api;
}

/** Internal — used by `<UniverSheet>` to publish the API once ready. */
export function useSetUniverAPI(): (api: FUniver | null) => void {
  const ctx = useContext(UniverContext);
  if (!ctx) {
    throw new Error('useSetUniverAPI must be used inside <UniverProvider>');
  }
  return ctx.setApi;
}
