---
'@casualoffice/sheets': minor
---

Gate the Help menu and its branding links behind feature flags so an embedded
host can present the editor as fully native (no "View on GitHub" / About /
editor-branded surfaces).

- `features={{ help: false }}` drops the whole Help menu.
- `features={{ branding: false }}` drops the "View on GitHub" and "About casual
  sheets" links (from both the Help and File menus) while keeping Keyboard
  shortcuts in Help.

Additive and backward-compatible: with `features` unset, standalone chrome is
unchanged (GitHub + About still show). The pure menu-gating engine was
extracted to `chrome/menu-model.ts` so the contract is unit-tested.
