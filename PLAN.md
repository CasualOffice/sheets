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

### ✅ Phase 6 — Self-host platform (complete — v0.1.0)

The first version-bumped release. v0.0.x was the "build a real editor end-to-end" arc; v0.1.0 earns its self-host story.

- **WOPI host integration** — `host.Integration` TypeScript interface + 4 concrete backends behind `CASUAL_STORAGE`: `memory` (default, preserves v0.0.x shape) · `local` (filesystem) · `s3` (AWS / MinIO / R2 / B2) · `postgres` (single bytea table).
- **WOPI endpoints**: CheckFileInfo · GetFile · PutFile (with `X-WOPI-ItemVersion` honoured as If-Match → 409 on mismatch).
- **JWT auth** — `CASUAL_JWT_SECRET` enables. Claims model: `sub` · `file_id` · `role` (admin/editor/commenter/viewer) · per-flag `permissions` · `features` toggles · `password_required` · `display_name`. URL `:id` must match `file_id` claim → 403 on cross-file lateral. `POST /api/tokens` admin-gated mint endpoint; `GET /api/me` self-introspection.
- **Admin panel** at `/admin` — env-gated by `CASUAL_ADMIN_USERNAME` + `CASUAL_ADMIN_PASSWORD`. Seven sections: branding · base path · storage (with per-backend cred forms) · networking (CORS / trust-proxy / HSTS / public origin) · room limits · auth providers (JWT live; OIDC + SAML stubbed for v0.2) · webhooks. JSON config persisted with mode 0600; secrets redacted on read (`***` sentinel preserves prior verbatim on write-back).
- **Webhook dispatcher** — 9 events (`room.created` / `dropped`, `file.uploaded` / `saved` / `deleted`, `user.joined` / `left`, `admin.login` / `login_failed`). HMAC-SHA256 signed via `X-Casual-Signature: sha256=<hex>` when subscription has a secret. Single retry after 5 s; v0.2 ships proper queue + dead-letter.
- **P6.1 — complex pivot cache passthrough** — `xl/pivotCaches/**` + `xl/pivotTables/**` survive round-trip. Rel renumbering across `xl/_rels/workbook.xml.rels` + every `xl/worksheets/_rels/sheet*.xml.rels`; `<pivotCaches>` injected into the ExcelJS-regenerated `xl/workbook.xml`. **Audit: 46/46 → 54/54**.
- **OCI image labels** — `org.opencontainers.image.*` baked in from CI build args. Rolling-tag scheme: `0.1.0` · `0.1` · `0` · `latest`. SBOM + provenance attestations in the manifest.
- **Self-hosting + customization docs** — 11 new pages on schnsrw.live/docs/sheets/ covering overview · reverse-proxy recipes · TLS · CORS · scaling · backups · admin walkthrough · auth claims + token issuance · webhook signature verification (Node/Python/Go).
- **Mobile lane** (back-ported) — touch-pan driver synthesizes wheel events from pointermove (Univer 0.24 has no native touch-pan); compact chrome at ≤ 720 px / ≤ 480 px; sticky bottom action bar; formula bar input pinned to 16 px (iOS focus-zoom guard).
- **Test coverage** — unit tests **8 → 60** (host contract + WOPI routes + JWT auth matrix + admin config + webhook HMAC). E2E suite 357 + the home + mobile + audit specs.

---

### ✅ Phase 7 — Personal mode + WOPI maturation (complete — v0.3.1)

The "self-host platform" arc continued. v0.2 added personal accounts; v0.3 finished the WOPI bring-up + the IA refresh.

- **Phase C — Personal mode (single + multi)** — `UserStore` + bcrypt + SQLite at `<root>/.casual/users.db`; HMAC-signed session token, 30-day TTL. Per-user file scoping (`local.PerUserStores` → `<root>/users/<userID>/`). Full per-user CRUD (`POST/GET/PUT/PATCH/DELETE /files{/<id>}`). `Profile` sidecar with `displayName / timezone / locale / avatarUrl / prefs`. `casual-docs` CLI (`reset-password / list-users / promote / demote`) + admin routes behind `RequireAdmin`. First signup auto-promotes; structured logs + request-id middleware throughout. UI: `PersonalAuthGate`, `UserMenu`, `ProfileSettingsDialog`. See `docs/self-hosting/personal-mode.md`.
- **Phase D — WOPI (Mode 2)** — JWT verifier with JWKS cache (`alg`-confusion defence rejects HS\*). `GET /wopi/host` embed redirect; access_token threaded through the WS preflight. `host.Locker` capability with Lock/Unlock/RefreshLock; the room manager claims the host lock on first join, releases on drain. Per-room `RefreshLock` ticker (10 min default) so long sessions don't lose the host-side lock idle-out. `docID = base64url(wopiSrc)` keeps the gateway stateless.
- **M2 — Snapshot pipeline (client-push)** — `DocxEditorRef.save()` produces serialized bytes client-side; `useFileSourceAutoSave` pushes through `FileSource.save()` on a schedule. The "server-side Bun worker pool" originally tracked in M2 is deferred — client push covers the practical case without adding a Bun runtime to the production Docker image. `AutosaveStatus` gives the host a Google-Docs-style indicator.
- **Phase E — UX audit wave (2026-06-11/12)** — IA refresh + path router (`/home`, `/sheet/<id>`, `/sheet/new`), `MySpreadsheetsList` file picker (dedup by name, mobile-responsive with always-visible Delete on touch), URL rebinding on first save, empty-draft skip via `<EditTracker>`, AccountMenu Admin entry, `/r/<roomId>` auth-gate exemption, `document.title` route-driven, logout dirty-check, AccountMenu on /home, keyboard shortcuts cheat sheet + `formatShortcut` util (Mac sees ⌘, Win/Linux sees Ctrl), SaveStatusPill (Google-Docs "Saved X ago"), ActivityPill (persistent error log), Ctrl+Shift+P command-palette alias, collab display-name pre-fill from signed-in user. See `docs/UX_AUDIT.md` §5 — every item carries the SHA that shipped it.

## What's next

### P8 — Sharing model implementation (designed, not coded)

Design lives in `docs/SHARING_MODEL.md`. Phasing:

- **§6.1 Link tokens (single + multi)** — `POST /share/link` + token role enforcement in the join handshake + Link tab in share dialog. ~1–2 weeks.
- **§6.2 Member ACLs (multi only)** — routes + Members tab + email lookup. ~1 week.
- **§6.3 Audit log surface** — structured logs + admin-side log view. ~3 days.
- **§6.4 Suggestion mode** — separate proposal, TBD.

### P9 — Auth backends (carried from prior P7 plan)

- **OIDC + SAML backend** — UI shipped in v0.1, schema persists; enforcement still pending.
- **Webhook retry queue** — proper exponential backoff + dead-letter store.

### P10 — Scale-out (carried from prior P8 plan)

- Op-log scale for multi-hour rooms (current Stage-6 compaction is fine for typical sessions; bound growth more aggressively when we see real pressure).
- Horizontal scale-out beyond sticky sessions: stateless WebSocket via shared Yjs state, multi-region replication.

### Univer fork perf (separate pipeline)

See [`docs/UNIVER_FORK_PERF.md`](./docs/UNIVER_FORK_PERF.md). Items 5–10 (incremental scroll, sparse insert, lazy bootstrap, listener batching, formula coalescing, incremental dep-tree) are gated behind "Excel parity first" per [project memory](../../.claude/projects/-Users-sachin-Desktop-melp-services-sheet/memory/MEMORY.md) — finish app-side polish before the fork dive.

---

## Architecture reference

See [`docs/ARCHITECTURE.md`](./docs/ARCHITECTURE.md) for the full diagram and data flows.

| Concern | Pick |
| --- | --- |
| Grid + formula engine | Univer OSS 0.24.x |
| Frontend | React 18 + Vite + TypeScript strict |
| Collab | Yjs + Hocuspocus over WebSocket |
| xlsx I/O | ExcelJS + JSZip in Web Workers (with OOXML passthrough for VBA + pivots) |
| ods / csv / tsv | `@e965/xlsx` in Web Workers |
| Charts | ECharts overlay anchored to cell ranges |
| Workbook persistence | `host.Integration` interface (memory · local · S3 · Postgres) |
| Room persistence | Redis optional, 7-day TTL |
| Auth | JWT (HS256) with role + permission + feature claims |
| Admin | `/admin` React panel; env-gated; on-disk JSON config |
| Webhooks | 9 events with HMAC-SHA256 signing |
| Container | Node 22 Alpine, multi-arch (amd64 + arm64), OCI-labelled |
| Tests | 60 unit + 357 + home + mobile + lossiness Playwright e2e |
