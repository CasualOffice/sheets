---
'@casualoffice/sheets': minor
---

feat(sdk): `onBeforeCreateUnit` hook to register extra Univer plugins

`<CasualSheets onBeforeCreateUnit={(univer) => univer.registerPlugin(...)}>` fires
after the SDK registers its built-in plugins but before the workbook unit is
created — the only point at which a host can add register-time plugins (off-main
formula worker via `UniverRPCMainThreadPlugin`, crosshair-highlight, zen-editor,
…). Enables a power host to share the SDK editor core while keeping its own extra
plugins (Phase 3). NOT semver-covered — it hands over the raw `Univer` instance.
