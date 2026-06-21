---
'@casualoffice/sheets': minor
---

The iframe embed now ships the **full feature set** — tables, sort, filter, conditional formatting, data validation, drawing/images, hyperlinks, notes, thread comments, find/replace — matching the real app. Previously the embed ran `lazyPlugins={false}` (the minimal editor) to stay a single file. But the embed's tsup build is `splitting:false` + `noExternal:/.*/`, so the lazy loader's dynamic `import()`s are **inlined** into the one `embed-runtime.js` rather than emitted as chunks — the single-file deploy is preserved. Enabling lazy plugins means the embed eager-loads any feature whose data is already in the opened file (so tables/CF are never silently dropped) and idle-loads the rest, so the toolbar/menu feature actions (Insert ▸ Table, Data ▸ Filter, …) resolve.
