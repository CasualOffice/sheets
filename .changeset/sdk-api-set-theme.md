---
"@casualoffice/sheets": minor
---

`CasualSheetsAPI.setTheme('light' | 'dark')` — imperative light/dark switch, the
API equivalent of the reactive `appearance` prop. Flips Univer's
`ThemeService.setDarkMode` (canvas colours + the `univer-dark` class Univer
applies to the document root) via `api.setTheme(...)`, for hosts that drive the
editor through the ref rather than re-rendering with a prop.
