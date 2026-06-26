---
'@sheet/web': patch
---

Fix File → New overwriting the previously-open file in the desktop app. New replaced the workbook in-window but left the bridge's bound file path pointing at the old file, so the next Save wrote the empty workbook over it. New now clears the bound path so Save prompts for a location.
