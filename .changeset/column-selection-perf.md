---
'@sheet/web': patch
---

Fix multi-hundred-millisecond lag when selecting a whole column or row. The status-bar Sum/Avg/Count recompute scanned the selection's full nominal extent (up to the sheet's 1,048,576 rows) synchronously on the main thread; it now bounds the scan to the used range (last row/column with content), which is exact for all statistics and matches how Excel computes column stats instantly. Fixed in the Univer fork (`status-bar.controller.ts`).
