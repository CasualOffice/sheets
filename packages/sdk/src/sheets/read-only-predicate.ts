/**
 * Pure command-id predicates for the read-only / comment-only vetoes. Kept
 * Univer-free (no `@univerjs/*` imports) so the security boundary is unit-testable
 * under the bare `node --import tsx` runner — `read-only.ts` pulls in Univer
 * services and can't load there.
 */

/**
 * Command ids that MUTATE a sheet — opening the cell editor, writing values,
 * styling, structural edits, clipboard paste. The read-only veto cancels any
 * command whose id matches. Navigation (selection, scroll, zoom, sheet switch),
 * copy, and undo/redo deliberately fall through so preview stays usable.
 */
export const READONLY_BLOCK =
  /(set-cell-edit-visible|set-activate-cell-edit|set-range-values|set-style|set-bold|set-italic|set-underline|set-strike|set-font|set-background|set-text|set-horizontal|set-vertical|set-wrap|set-rotation|set-border|set-number-format|insert-|delete-|remove-|clear-selection|cut-content|paste|move-range|move-rows|move-cols|merge|split|add-worksheet|set-worksheet-name|set-worksheet-row|set-worksheet-col|auto-fill|reorder|set-defined-name|set-tab-color|set-frozen-cancel)/;

/**
 * Commands that must stay usable for the **comment** share-role even though the
 * cells are locked: the threaded-comment commands (add / reply / update /
 * resolve / delete) and the comment editor's rich-text doc edits. Some of these
 * (`delete-comment`) would otherwise be caught by {@link READONLY_BLOCK}'s
 * `delete-` token, so the comment-only veto allowlists them explicitly.
 */
export const COMMENT_ALLOW = /(thread-comment|rich-text-editing)/;

/** True for any mutating command (the read-only veto target). */
export function isReadOnlyBlocked(commandId: string): boolean {
  return READONLY_BLOCK.test(commandId);
}

/**
 * The comment-only veto predicate (the security boundary): block cell/style/
 * structural mutations, but let threaded-comment + comment-editor commands
 * through.
 */
export function isCommentOnlyBlocked(commandId: string): boolean {
  return READONLY_BLOCK.test(commandId) && !COMMENT_ALLOW.test(commandId);
}
