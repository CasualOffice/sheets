---
'@casualoffice/sheets': minor
---

Make the PivotTable feature work end-to-end, and fix persistence of the SDK's custom snapshot resources.

- **Custom-resource persistence (the foundation):** the charts / pivots / sparklines panels persist their models as `IWorkbookData` resources, but Univer silently drops any resource it doesn't own on `getContent()`/`setContent()` — so those models never round-tripped (no save/reload, no collab). The SDK now shadows these resources in an api-layer store: `setContent` captures them before the workbook swap and `getContent` re-attaches them. Existing panel/chart code is unchanged; it now just persists.
- **Pivot end-to-end:** the pivot engine (`computePivot`/`applyPivot`) is ported into the SDK. The Insert PivotTable dialog now creates + persists an editable model (not just a static grid), the Fields panel picks it up, and every field edit re-applies via the engine so the output grid recomputes live — no more "config saved but output stale".
