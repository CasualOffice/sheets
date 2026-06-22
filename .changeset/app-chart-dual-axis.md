---
'@sheet/web': minor
---

Add **combo charts + a secondary (dual) value axis** (chart Format dialog ▸ "Series type & axis"). Per-series controls let a column / line / area chart mix bars and lines (Excel's Combo chart type) and plot any series against a secondary right-hand value axis — the chart then renders two `yAxis` entries with the flagged series routed to `yAxisIndex: 1`. Overrides persist on the chart's `format` (so they survive reload + xlsx round-trip), default to the chart's base type/axis, and are gated to the families where they read correctly (not pie / doughnut / scatter / 100%-stacked / horizontal bar). Closes a common gap vs Excel for revenue-vs-margin style charts.
