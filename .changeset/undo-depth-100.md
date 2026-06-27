---
'@casualoffice/sheets': patch
---

Raise the undo depth from 20 to 100 to match Excel. The Univer fork capped the
undo stack at 20 levels — a frequent power-user papercut, since a handful of
find-and-replaces or fill-downs would exhaust it and silently drop older
history. Each undo entry holds only mutation params (range refs + values), so
the deeper stack is cheap for the common edit.
