---
'@casualoffice/sheets': minor
---

feat(chrome): zoom control in the status bar

The built-in chrome status bar gains a zoom control on the right: − / level / +,
with the level click resetting to 100%. Dispatches `sheet.operation.set-zoom-ratio`
(clamped 10–400%) — the operation path, since the higher-level zoom commands bail
when Univer's formula-bar editor unit reports visible. Closes the zoom gap that an
earlier batch deferred (the block was a test-timing artifact, not a real
registration problem — Univer's zoom render controller registers in `onRendered`).
