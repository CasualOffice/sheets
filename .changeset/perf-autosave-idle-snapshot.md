---
'@sheet/web': patch
---

Stop autosave from hitching the grid on large workbooks. The autosave snapshot is a full `wb.save()` deep clone that can take hundreds of milliseconds on a big sheet, and the 30-second tick ran it on the main thread regardless of what the user was doing — freezing the grid mid-keystroke. The snapshot now runs in a `requestIdleCallback` slot (with a 2s timeout guarantee, and a `setTimeout` fallback for older Safari), so it lands when the browser is idle rather than mid-edit. The `pagehide`/`beforeunload` flush stays synchronous, and the saved content is unchanged — only the timing moves.
