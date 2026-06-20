---
'@sheet/web': minor
---

apps/web shares the SDK editor core (Phase 3 step 1)

`apps/web` no longer hand-rolls its Univer bootstrap — `UniverSheet.tsx` now
renders `<CasualSheets chrome="none">` from `@casualoffice/sheets`, sharing the
SDK's Univer boot, plugin set, formula engine, and snapshot/API. The app keeps
its rich shell (ribbon, charts, pivots, panels, dialogs) and layers its extras
on top: crosshair-highlight + zen-editor + Merge/Unmerge context menu via
`onBeforeCreateUnit`, off-main compute via `formula={{ worker }}`, and the
paste-merge hook / dev helpers / zoom-shortcut override via `onReady`. One Univer
bootstrap now serves both the app and third-party SDK hosts.
