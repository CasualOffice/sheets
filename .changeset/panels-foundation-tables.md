---
'@casualoffice/sheets': minor
---

Add a side-panel rail to the SDK chrome (`chrome:"full"`/`"embedded"`), so embedders get the panels the standalone app has instead of a bare grid. This first drop ships the rail + a one-at-a-time panel store (mutex) + a built-in-panel registry, and the **Tables** panel (list / rename / re-theme / delete the tables on the active sheet). Host panels registered via `extensions.panels` appear on the same rail. Charts, Pivot, Comments and History panels follow.
