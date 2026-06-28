---
'@casualoffice/sheets': minor
'@sheet/web': minor
---

Native pivot export (opt-in). The SDK now exports `generateNativePivot` (build real `xl/pivotTables` + `xl/pivotCaches` OOXML from a pivot model) and `applyPivotsToZip`, so a host can compose native PivotTables into an export. The app wires this behind an off-by-default flag (`cs-native-pivots`): when enabled, in-app pivots round-trip to Excel as real, refreshable PivotTables instead of flat cells. Default behaviour is unchanged.
