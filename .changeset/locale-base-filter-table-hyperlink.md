---
'@sheet/web': patch
'@casualoffice/sheets': patch
---

Fix raw i18n keys in the filter, table and hyperlink features. Like the data-validation fix in #252, the locale bundle merged only the `-ui` halves of these features, so their error toasts and generated labels rendered raw keys (e.g. `sheets-table.tablePrefix`, `sheets-filter.command.not-valid-filter-range`, `sheets-hyper-link.message.refError`). The base `@univerjs/sheets-filter`, `@univerjs/sheets-table` and `@univerjs/sheets-hyper-link` locales are now merged into both the app bundle and the SDK embed runtime. (All three are already pinned + fork-linked deps, so no dependency changes.)
