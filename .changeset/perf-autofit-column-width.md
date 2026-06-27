---
'@casualoffice/sheets': patch
---

Speed up auto-fit column width on large sheets. The Univer fork's
`calculateAutoWidthInRange` built a full `DocumentViewModel` +
`DocumentSkeleton` and laid it out for every measured cell — on a whole-sheet
auto-fit that is tens of thousands of layouts, so fitting a 21k-row × 8-col
sheet froze the UI for ~10s. The common cell (plain value, no rich text, no
wrap, no rotation) now measures its widest line with the LRU-cached
`FontCache` — the same primitive the renderer uses to size non-wrap content —
and falls through to the old `DocumentSkeleton` path only for wrap / rotation /
rich-text cells. Auto-fitting that 21k × 8 sheet drops from ~10s to ~0.3s, with
pixel-consistent widths. Validated by the engine-render integration suite plus
a new end-to-end auto-fit benchmark.
