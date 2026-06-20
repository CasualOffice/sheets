---
'@casualoffice/sheets': minor
---

feat(sdk): `formula={{ worker }}` for off-main formula compute

By default `<CasualSheets>` computes formulas on the main thread (fine for typical
sheets, zero host setup). Pass a Web Worker to move compute off-thread so paste /
sort / fill on large workbooks don't freeze the UI: the SDK registers the formula
plugins with `notExecuteFormula` and wires `UniverRPCMainThreadPlugin` to your
worker (dynamic-imported, so `@univerjs/rpc` stays a true optional peer — only
loaded when a worker is passed; the host owns the worker, the SDK never bundles
one). This is the second enabler (with `onBeforeCreateUnit`) for a power host to
share the SDK editor core without regressing off-main compute.
