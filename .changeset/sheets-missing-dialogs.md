---
'@casualoffice/sheets': minor
---

Add the 12 missing built-in dialogs so their menu items do something instead of opening nothing: Data Validation, Conditional Formatting, Sort range, Paste Special, Insert Function, Name Manager, Insert cells, Delete cells, Goal Seek, Insert Chart, Insert Sparkline, and Insert Pivot. Each is a real form wired to the Univer facade (data-validation / conditional-formatting / sort / clipboard / defined-names / range commands) and registered in the chrome's built-in dialog set, so embedders get a functional editor rather than inert menu entries.
