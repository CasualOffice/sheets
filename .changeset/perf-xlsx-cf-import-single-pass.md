---
'@casualoffice/sheets': patch
---

Speed up xlsx import. The raw-OOXML conditional-formatting captures (data-bar colours + duplicate/unique rules) were each loading the zip and re-decompressing every worksheet's XML separately — adding ~455ms to parsing a 160k-cell workbook even when it had no conditional formatting. They're now merged into a single zip pass that decompresses each worksheet once, cutting large-file parse time ~40% (≈994ms → ≈590ms in that benchmark). Behavior is unchanged.
