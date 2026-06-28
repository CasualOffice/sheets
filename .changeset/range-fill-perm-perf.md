---
'@sheet/web': patch
---

Avoid a synchronous full-range permission scan when dragging the fill handle from a very large (e.g. whole-column) selection on a sheet with no protection rules. The fill-handle permission check now short-circuits when the fill range overlaps no protected range — matching the range-move check — instead of walking every cell. Fixed in the Univer fork.
