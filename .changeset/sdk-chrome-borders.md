---
'@casualoffice/sheets': minor
---

feat(chrome): borders dropdown in the toolbar

The built-in chrome toolbar gains a borders control (next to the colour pickers):
a dropdown with All / Outside / Inside / Top / Bottom / Left / Right / No border.
Each dispatches `sheet.command.set-border-position` against the active selection
using Univer's current border style/colour. Closes a common formatting gap
between the SDK chrome and a real spreadsheet editor.
