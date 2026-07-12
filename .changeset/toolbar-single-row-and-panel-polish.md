---
'@casualoffice/sheets': patch
---

Fix the toolbar wrapping into a broken second row and unify the side-panel UI. The formatting toolbar now stays a single row and scrolls horizontally when the controls don't fit (like Google Sheets), instead of wrapping and stranding icons on a second line. All side panels (Tables, Charts, Pivot, Comments, History) now share one polished shell — a consistent header (icon + title + count + hover-capable close) and empty-state — replacing per-panel inline markup; the Charts panel in particular was relying on CSS classes that don't ship in the SDK, so it rendered unstyled. Styling maps to the design-system `--color-*` tokens (with chrome-var fallbacks) so it themes correctly inside a host.
