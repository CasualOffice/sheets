---
'@casualoffice/sheets': minor
---

Preserve the remaining text conditional-formatting operators — `beginsWith`, `endsWith`, and `notContainsText` — through the xlsx round-trip (on top of the existing `containsText` and blanks/errors predicates). ExcelJS surfaces these rules fully on read (type, operator, formula, style) and writes them back when given an explicit formula, so the search string is recovered from / written into the rule formula and they round-trip and paint on open. With this, every Excel conditional-formatting rule type now round-trips.
