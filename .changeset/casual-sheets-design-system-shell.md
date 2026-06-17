---
'@schnsrw/casual-sheets': minor
---

Ship `/shell` subpath export and adopt `@schnsrw/design-system` tokens in
the embed runtime.

- **New `/shell` entry** — TitleBar, Toolbar, FormulaBar, SheetTabs,
  StatusBar, and a cohesive `SheetShell` wrapper, ported from the Casual
  Office design bundle and built on `@schnsrw/design-system` primitives.
  Props-driven components that consumers compose with `<CasualSheets>`
  (the embed runtime and Drive both use them).
- **Embed runtime imports `@schnsrw/design-system/tokens.css`** so the
  iframe paints in the canonical token vocabulary (Inter + JetBrains Mono
  - Manrope + Material Symbols Outlined; teal accent for Sheets, cyan for
    the docs editor variant).
- **Theme command wired** — the runtime now subscribes to
  `casual.command.set.theme` and flips `data-theme` on `<html>` to
  `light` / `dark`, or clears it for `system`. Hosts (like Drive) can
  drive iframe theming via the protocol without resorting to the
  copy-embed.mjs MutationObserver hack.
- **`data-app` set from URL** — the iframe applies the docs editor's
  cyan accent ramp when `?app=docs`, the sheet's teal ramp otherwise.

No breaking changes. Existing `casual.command.execute` and
`casual.selection.format-state` envelopes remain — they are now marked
`@experimental` because the polished chrome lives in the iframe under
`/shell` and headless hosts that drive Univer through the wire are no
longer the recommended path for new integrations.
