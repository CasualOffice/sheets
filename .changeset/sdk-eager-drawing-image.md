---
'@casualoffice/sheets': patch
---

Fix "can't insert image". Insert ▸ Image opens a file picker, which needs the user's click gesture — but the drawing plugin was lazy/idle-loaded, so a quick click before it registered silently no-opped, and `await`-ing to load it on click lost the gesture (picker never opened). The drawing plugin now eager-loads during boot (tracked, so idle-load doesn't double-register), so the image picker opens on the first click. Other features open panels (no gesture) and are unaffected.
