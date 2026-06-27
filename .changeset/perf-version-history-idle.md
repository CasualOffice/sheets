---
'@sheet/web': patch
---

Defer the version-history auto-snapshot to an idle slot too, matching the autosave fix. Its ~10-minute capture interval ran a full `wb.save()` deep clone on the main thread regardless of activity, which could freeze the grid mid-edit on a large workbook. The shared `runWhenIdle` helper (extracted to `idle.ts`, now used by both the autosave and version-history loops) runs the clone when the browser is idle instead.
