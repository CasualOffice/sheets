import type { Peer } from './presence';

export type UserEntry = { userID: string; name: string };

/**
 * Build the (current-user, peer-users) set to push into Univer's
 * UserManagerService. Pure + Univer-free so the dedup/shape logic is
 * unit-testable. Drops peers with no stable userId and de-dupes by userID
 * (self wins; first peer wins on collision).
 */
export function buildCollabUserSet(
  self: UserEntry,
  peers: Pick<Peer, 'userId' | 'name'>[],
): { current: UserEntry; others: UserEntry[] } {
  const seen = new Set<string>([self.userID]);
  const others: UserEntry[] = [];
  for (const p of peers) {
    if (!p.userId || seen.has(p.userId)) continue;
    seen.add(p.userId);
    others.push({ userID: p.userId, name: p.name });
  }
  return { current: self, others };
}
