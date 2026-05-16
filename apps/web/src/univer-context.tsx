import { createContext } from 'react';
import type { FUniver } from '@univerjs/core/facade';

export type UniverCtxValue = {
  api: FUniver | null;
  setApi: (api: FUniver | null) => void;
};

export const UniverContext = createContext<UniverCtxValue | null>(null);
