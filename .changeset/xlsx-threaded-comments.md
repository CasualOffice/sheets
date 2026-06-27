---
'@casualoffice/sheets': patch
---

Preserve threaded comments — authors, timestamps, and reply chains — across an
xlsx round-trip. Modern Excel comments live in `xl/threadedComments/**` +
`xl/persons/**`, but ExcelJS only models the legacy note, so our bridge
collapsed every thread to a single note authored "imported" and dropped the
replies/authors on save. A new threaded-comment passthrough captures the
threaded layer at parse time and re-injects it at export — restoring the parts,
declaring their content types, and re-creating the workbook→persons and
sheet→threadedComment relationships (the parts are discovered by relationship
type, so no XML-element injection is needed). It rides on top of the legacy
note ExcelJS still writes, so for an open → save round-trip the full
author/reply metadata survives in Excel.
