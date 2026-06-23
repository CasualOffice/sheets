import type { FUniver } from '@univerjs/core/facade';
import { UserManagerService } from '@univerjs/core';
import type { Peer } from './presence';
import { buildCollabUserSet, type UserEntry } from './user-manager-sync.pure';

export { buildCollabUserSet } from './user-manager-sync.pure';

interface UserManagerLike {
  setCurrentUser(u: UserEntry): void;
  addUser(u: UserEntry): void;
}

/**
 * Populate Univer's UserManagerService from the collab session so comment
 * authorship (`personId`) and @mention candidates resolve to display names.
 * Without this the service is empty and comment authors render as the default
 * user. Idempotent (UserManagerService stores users in a Map keyed by userID).
 */
export function syncCollabUsers(api: FUniver, self: UserEntry, peers: Peer[]): void {
  const injector = (api as unknown as { _injector?: { get(t: unknown): unknown } })._injector;
  const svc = injector?.get(UserManagerService) as UserManagerLike | undefined;
  if (!svc) return;
  const { current, others } = buildCollabUserSet(self, peers);
  try {
    svc.setCurrentUser(current);
    for (const u of others) svc.addUser(u);
  } catch {
    /* best-effort — author names are a nicety, never block the session */
  }
}
