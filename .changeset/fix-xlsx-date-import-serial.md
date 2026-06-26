---
'@casualoffice/sheets': patch
---

Fix xlsx import storing date cells as ISO strings instead of Excel serial numbers. ExcelJS surfaces date/time-formatted cells as JS Dates; the importer wrote them as `toISOString()`, so date functions (e.g. `=NETWORKDAYS(E8,F8)`) couldn't parse their operands and errored, and dates displayed as raw ISO text. Date cells now import as Excel serial numbers (preserving the date number-format), so date math evaluates and the cell renders and round-trips as a date.
