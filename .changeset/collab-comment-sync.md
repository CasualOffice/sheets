---
'@sheet/web': minor
---

Real-time comment sync across co-editors. Thread comments (add / edit / resolve / delete / re-anchor) now cross the Yjs op-log bridge so a comment created, resolved, or deleted by one collaborator shows up live for everyone in the room — previously they stayed local-only. The five `thread-comment.mutation.*` ids are added to the bridge's `SYNCED_MUTATIONS` allowlist and mapped to the `threadComment` lazy-plugin group so a peer that hasn't opened the Comments pane loads the plugin before replaying the change (no silent drops). The `threadComment` lazy loader now also pulls `@univerjs/sheets-thread-comment/facade`, installing `FWorksheet.getComments` / `FRange.addComment` so the Comments pane and facade-driven flows work once the plugin mounts. Covered by a new two-client `coedit-comments` Playwright spec asserting bidirectional propagation.
