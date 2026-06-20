---
'@casualoffice/sheets': minor
---

feat(chrome): richer status-bar stats (Numerical Count / Min / Max)

The built-in chrome status bar now shows Excel's full selection-aggregate set —
Average, Count, Numerical Count, Min, Max, Sum. Count is non-empty cells (any
type); Numerical Count is numeric cells; the numeric aggregates run over numeric
cells only, matching Excel. (A zoom control is deferred to a follow-up batch: the
SDK's eager plugin set doesn't yet register Univer's zoom render controller.)
