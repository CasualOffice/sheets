import { createContext, useContext } from 'react';

export type CollabStatus = 'off' | 'connecting' | 'live' | 'offline' | 'denied';
export type CollabRole = 'view' | 'write';

export type CollabCtxValue = {
  /** True when the active build was made with VITE_COLLAB_ENABLED. */
  enabled: boolean;
  /** Current room id, or null when not in a room. */
  roomId: string | null;
  /** WebSocket transport status. `off` when no room is joined; `denied`
   *  when the server rejected the password / auth. */
  status: CollabStatus;
  /** Effective role for this client in the active room. Defaults to
   *  `write` outside a room (the local single-user editor). */
  role: CollabRole;
};

export const CollabContext = createContext<CollabCtxValue>({
  enabled: false,
  roomId: null,
  status: 'off',
  role: 'write',
});

export function useCollab(): CollabCtxValue {
  return useContext(CollabContext);
}
