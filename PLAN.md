# Casual Sheets — Plan

Web-based Excel-equivalent with real-time co-editing.
Upload `.xlsx` → open in browser → multiple users edit together, anonymously, no accounts.

---

## Scope

### In scope
- Excel / Office UX: ribbon, formula bar, file-centric workflow.
- Upload `.xlsx` → open in a shared room → download back as `.xlsx`.
- Multi-user co-editing in real time (cursors, selections, live typing).
- In-memory sessions — state lives while at least one user is connected, with a grace-period GC after the last client leaves.
- Optional Redis persistence for sessions that survive server restarts.

### Out of scope (deferred)
- **Persistence / WOPI** — no DB, no autosave to disk, no SharePoint/OneDrive integration.
- **Auth / accounts** — anonymous sessions identified by room URL.
- **AI / LLM features** — the Univer command bus is extensible; plug your own model in later.
- **Mobile** — desktop browsers only.

### Out of scope (forever)
- 100% Excel feature parity — Excel has 30+ years of features; we ship the 80% that matter.
- Pixel-perfect Office clone — "clearly inspired by / familiar to Office users" is the bar.

---

## Phase status

### ✅ Phase 0 — Spikes (complete)
All three technical risks proved out:
- **Spike A (Yjs bridge)** — two browsers, edits visible in < 250 ms.
- **Spike B (xlsx round-trip)** — ExcelJS parse + export with acceptable fidelity.
- **Spike C (UI override)** — Univer default chrome hidden, custom ribbon wired.

### ✅ Phase 1 — Single-user Excel-flavored editor (complete)
- Custom Office-style ribbon: Home, Insert, Formulas, Data, View, Review tabs.
- Formula bar with editable Name Box.
- Status bar: SUM / AVG / COUNT / MIN / MAX, zoom, sheet tabs.
- Full xlsx / ods / csv / tsv open + save, worker-side.
- All major Excel keyboard shortcuts.
- 337 Playwright e2e tests locking the surface.

### ✅ Phase 2 — Real-time co-editing (complete)
- Hocuspocus server + Yjs bridge plugin.
- Room lifecycle: create on upload, GC after TTL.
- Share URL: anyone joins. Password-protected rooms with SHA-256 gate.
- View-only role enforced at the Univer engine layer.
- Redis persistence (optional, 7-day TTL).
- Joiner fast-path: gzip-streamed snapshot, skips xlsx parse.
- Op-log compaction (Stage 6) for long-lived rooms.
- Self-hosted Docker image: web + Hocuspocus + Fastify in one container.

### ✅ Phase 3 — Presence + polish (complete)
- Peer cursors on the grid (scroll-tracked, frozen-pane-aware, zoom-aware).
- Live-typing ghost: characters appear in the peer's cell as they type.
- Presence avatars with "Active now / Last seen Ns ago" tooltips.
- "Waiting to reconnect" banner + faster offline detection.
- Divergence detector: amber "Out of sync" pill when state vectors diverge.
- Session-history side panel: per-room op log, timestamps, revert.

### ✅ Phase 4 — Feature breadth + co-edit fidelity (complete — v0.0.5)
- Charts P1–P5b: insert dialog, 8 chart types, drag/resize, format dialog, collab sync, PNG embed in xlsx.
- Pivot tables P0: group-by + aggregate from Insert menu.
- Conditional formatting, data validation, drawings — all co-edit synced.
- Workbook/worksheet metadata (tab colors, zoom, freeze, sheet visibility) co-edit synced.
- Autosave to IndexedDB with restore banner on reload.
- ODS fidelity: styles, dimensions, freeze, hyperlinks, comments, defined names.
- 30+ additional Excel keyboard shortcuts.

### ✅ Phase 5 — Excel-parity wave (complete — v0.0.6)
- **Analysis tools**: Name Manager (Ctrl+F3), Flash Fill (Ctrl+E), Goal Seek (iterative solver dialog).
- **Pivot tables P1**: filter fields, Refresh PivotTables, drill-down to contributing source rows (Ctrl+Shift+D).
- **Charts**: trendlines, date-axis detection, per-series colour overrides.
- **Sparklines**: in-cell mini-charts (line / column / win-loss), workbook resource persistence for xlsx round-trip.
- **Show Formulas** (Ctrl+`): non-destructive DOM overlay that paints formula source on every formula cell.
- **Print Area**: A1 field in Page Setup + File-menu "Set / Clear Print Area" against the active selection.
- **Recent Files landing**: IndexedDB-backed, last-10 entries, surfaced when the workbook is a blank `Untitled`.
- **Paste Special** dialog (Ctrl+Alt+V): 6 Univer-native paste modes wired through the standard mutation path.
- **Server-side view-only enforcement**: Hocuspocus `onAuthenticate` flips `connection.readOnly` so view-role joiners can't bypass the client gate.
- **Theme**: dark mode toggle (title-bar sun/moon), bridged to Univer's `ThemeService` so the canvas chrome flips too.
- **Shell rewrites**: Google-Docs-style merged title bar (logo + filename + menus + actions), right-edge panel rail (Tables / Charts / Outline / Comments / History), brand mark aligned with sister doc-editor.
- **Inline SVG icons**: ~155 components covering every name the app uses; sharp at every size, no font-load delay.
- **Local version history**: snapshot store + preview + restore in single-user mode.
- **Status-bar customisation**: right-click checklist (Average / Count / Sum / Min / Max / Numerical Count).
- **Multi-range presence**: peer cursors render every range in a Ctrl-click selection.
- **NamePill**: in-room name edit affordance.
- **Quick wins**: `Ctrl+Alt+L` re-apply filter, `Ctrl+[ / Ctrl+]` precedent / dependent navigation.

---

## What's next

### P6 — Remaining gaps (small)

| Area | What's needed |
| --- | --- |
| Pivots | ~~Multi-row-field with Excel's compact layout~~ ✅ (P1.5 — outer subtotals + indented inner rows + grand total; sub-row dropdown on Insert dialog; drill-down follows the full composite key path) |
| xlsx round-trip | VBA stub passthrough ✅ (lossiness audit now 46/46 — .xlsm files round-trip macros byte-equal). Complex pivot cache deferred to **P6.1** (rel renumbering + workbook.xml surgery); numfmt long tail ✅ |
| UX | ~~Display-name edit surface post-join~~ ✅ (join prompt now carries a "Joining as" field so users can override their stored name BEFORE peers see them; NamePill stays the post-join surface) |

### P6.1 — Complex pivot cache passthrough (deferred)

xlsx files authored with pivot tables currently lose the `xl/pivotCaches/**`
and `xl/pivotTables/**` parts on round-trip. The clean fix is the same
byte-passthrough pattern shipped for VBA, but pivots need extra OOXML
surgery: rel renumbering across `xl/_rels/workbook.xml.rels` +
`xl/worksheets/_rels/sheet*.xml.rels` and injection of `<pivotCaches>`
into the ExcelJS-regenerated `xl/workbook.xml`. Estimated 2 days. Slot
in alongside the home/templates work as a side track.

### P7 — WOPI host integration (deferred)

When persistence is required: add a Hocuspocus persistence adapter (Postgres or S3-backed) and implement the WOPI host contract. The collab layer doesn't change — only the storage backend.

### P8 — Scale (deferred)

- Op-log scale for multi-hour rooms (current Stage-6 compaction is fine for typical sessions; bound growth more aggressively when we see real pressure).
- Horizontal scale-out: single-process today.

---

## Architecture reference

See [`docs/ARCHITECTURE.md`](./docs/ARCHITECTURE.md) for the full diagram and data flows.

| Concern | Pick |
| --- | --- |
| Grid + formula engine | Univer OSS 0.24.x |
| Frontend | React 18 + Vite + TypeScript strict |
| Collab | Yjs + Hocuspocus over WebSocket |
| xlsx I/O | ExcelJS in Web Workers |
| ods / csv / tsv | `@e965/xlsx` in Web Workers |
| Charts | ECharts overlay anchored to cell ranges |
| Persistence | Redis optional, 7-day TTL |
| Container | Node 22 Alpine, multi-arch (amd64 + arm64) |
| Tests | Playwright (Chromium), 337 e2e tests |
