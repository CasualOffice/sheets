---
'@sheet/web': minor
---

Add **PivotTable column fields** (cross-tab / matrix layout). The Insert ▸ PivotTable dialog gains a **Column field** dropdown: placing a field on the column axis fans the value field out across one column per distinct value, producing Excel's matrix layout (`Region` down the rows, `Quarter` across the columns) with a right-hand **Grand Total** column carrying row totals, a bottom Grand Total row carrying column totals, and the overall total in the corner. Works with multi-row compact layouts and filters; drill-down ("show details") now narrows by the clicked column's key as well as the row key. Picking the same field for both rows and columns falls back to the classic row-only pivot. Closes the column-field gap noted in the P0 pivot spec.
