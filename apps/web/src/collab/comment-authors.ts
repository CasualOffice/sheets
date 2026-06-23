/**
 * Comment authorship store — who wrote each comment.
 *
 * Univer's thread comments stamp `personId` from the *current user*
 * (`UserManagerService`). We can't set a distinct current user per client:
 * `setCurrentUser` feeds Univer's permission layer, and a non-default user id
 * flips the collab grid to read-only (the #122 regression). So every client
 * shares the empty default `personId` and comments carry no usable author.
 *
 * Instead we record authorship out-of-band, keyed by comment id, exactly like
 * charts ride their own Y.Map alongside the cell-mutation op-log:
 *   - the *creating* client stamps `commentId → {name,color}` from its own
 *     presence identity (the only client that runs the add-comment **command**;
 *     peers only receive the replicated mutation),
 *   - in a room that stamp is mirrored into a `casual-comment-authors` Y.Map so
 *     peers resolve the name,
 *   - single-player just keeps the local map (author = the local display name).
 *
 * This module is the transport-agnostic store. CollabDriver owns the stamping
 * hook + Y.Map mirror; CommentsPanel subscribes for render.
 */

export type CommentAuthor = { name: string; color: string };

const store = new Map<string, CommentAuthor>();
const listeners = new Set<() => void>();
let version = 0;

function emit(): void {
  version += 1;
  for (const l of listeners) l();
}

function same(a: CommentAuthor | undefined, b: CommentAuthor): boolean {
  return !!a && a.name === b.name && a.color === b.color;
}

/** Record one comment's author. Returns true if it was new/changed. */
export function recordCommentAuthor(id: string, author: CommentAuthor): boolean {
  if (same(store.get(id), author)) return false;
  store.set(id, author);
  emit();
  return true;
}

/** Merge a batch (e.g. a remote Y.Map snapshot). Emits once if anything changed. */
export function mergeCommentAuthors(entries: Iterable<[string, CommentAuthor]>): void {
  let changed = false;
  for (const [id, a] of entries) {
    if (!same(store.get(id), a)) {
      store.set(id, a);
      changed = true;
    }
  }
  if (changed) emit();
}

export function getCommentAuthor(id: string): CommentAuthor | undefined {
  return store.get(id);
}

/** Snapshot of every known author — used to seed a freshly-joined room's map. */
export function snapshotCommentAuthors(): Array<[string, CommentAuthor]> {
  return Array.from(store.entries());
}

/**
 * Subscribe for change notifications. Pair with {@link commentAuthorsVersion}
 * in `useSyncExternalStore` — the version is a stable snapshot value that
 * changes on every mutation, so React re-reads `getCommentAuthor` per row.
 */
export function subscribeCommentAuthors(fn: () => void): () => void {
  listeners.add(fn);
  return () => {
    listeners.delete(fn);
  };
}

export function commentAuthorsVersion(): number {
  return version;
}

/** Test-only: wipe the store between cases. */
export function __resetCommentAuthors(): void {
  store.clear();
  emit();
}
