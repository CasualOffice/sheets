---
'@casualoffice/sheets': minor
---

Add the Comments, PivotTable Fields, and History panels to the SDK chrome rail (following the Tables panel from the previous release), so embedders get them natively:

- **Comments** — lists thread comments on the active sheet (open + resolved), resolve/reopen, and add-comment; clicking a row navigates to the cell for full reply threading. Author avatars/mentions are host-owned and degrade to the in-band author.
- **PivotTable Fields** — edit a pivot's field layout (Rows/Columns/Values/Filters, aggregation, show-values-as, date grouping); the definition round-trips through the workbook. Recomputing the laid-out grid still needs the pivot engine (a later change) — the panel notes this.
- **History** — an activity feed of recent edits (from the command service). Persistent version history remains host-owned.
