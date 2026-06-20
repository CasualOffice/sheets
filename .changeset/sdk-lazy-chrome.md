---
'@casualoffice/sheets': minor
---

perf(chrome): lazy-load the built-in chrome (`chrome="none"` no longer bundles it)

`<CasualSheets>` now `lazy`-imports its chrome from the new `@casualoffice/sheets/chrome`
subpath only when `chrome !== 'none'`. The subpath is externalised in the build, so
the chrome stays a separate chunk the consumer's bundler code-splits — bare-grid
hosts (the default, and any `chrome="none"` integrator) no longer carry the chrome
JS. `dist/sheets.js` drops from ~62 KB to ~24 KB; the chrome ships as `dist/chrome.js`
loaded on demand. The bars now appear a tick after first paint (lazy chunk load).
