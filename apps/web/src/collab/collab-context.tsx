import { createContext, useContext } from 'react';

export type CollabStatus = 'off' | 'connecting' | 'live' | 'offline';

export type CollabCtxValue = {
  /** True when the active build was made with VITE_COLLAB_ENABLED. */
  enabled: boolean;
  /** Current room id, or null when not in a room. */
  roomId: string | null;
  /** WebSocket transport status. `off` when no room is joined. */
  status: CollabStatus;
};

export const CollabContext = createContext<CollabCtxValue>({
  enabled: false,
  roomId: null,
  status: 'off',
});

export function useCollab(): CollabCtxValue {
  return useContext(CollabContext);
}
