---
'@casualoffice/sheets': minor
---

Embed now honors the host's **light/dark theme**. Previously the iframe always rendered light (it never set `appearance`), so it didn't match a dark host. The runtime now reads `?theme=light|dark|system` from the embed URL, resolves `system` against the iframe's `prefers-color-scheme` (and follows live OS changes), and passes `appearance` to `<CasualSheets>` so Univer's canvas/headers/gridlines + the SDK chrome all theme together. Hosts can also push live changes over `casual.command.set.theme`.

Also fixes a protocol bug: `EmbedHostTransport` posted `casual.command.set.{theme,readonly,locale}` (dotted) but the runtime listened for `setTheme`/`setReadOnly`/`setLocale` (camelCase), so those three host→editor commands were silently dropped. Aligned both sides to the dotted form.
