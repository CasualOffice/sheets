---
'@casualoffice/sheets': patch
---

Preserve strikethrough and text rotation on xlsx import/export. Both were dropped by the style mapping: `font.strike` now round-trips through Univer's `st`, and angled or stacked (`vertical`) cell text round-trips through `tr` (preserving the angle and direction). Follows the border-style fidelity fix.
