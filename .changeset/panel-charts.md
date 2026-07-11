---
'@casualoffice/sheets': minor
---

Add the Charts panel and live chart overlay to the SDK chrome, so embedders can insert and edit charts natively (previously charts existed only in the standalone app). Insert a chart from the Charts panel (or Insert menu); it renders as an ECharts overlay anchored over the grid, moves/resizes, and persists in the workbook. echarts is lazy-loaded — it only downloads when a chart first renders or the Charts panel opens, so `chrome:"full"` hosts that never use charts don't pay for it. Column/bar/line/pie/scatter, combo/dual-axis, and trendlines are supported; chart definitions round-trip through the workbook snapshot.
