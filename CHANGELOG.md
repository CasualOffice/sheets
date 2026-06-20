# Changelog

All notable changes per release. Format loosely follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/) +
[SemVer](https://semver.org/).

Pre-v0.2.0 releases were tagged without a CHANGELOG file. The site
(`https://casualoffice.org/changelog/`) carries longer-form release notes
for v0.1.0+; the GitHub Releases page links to the same content.

## [0.3.3] — 2026-06-20

### Fixed

- **Personal mode crashed with `SqliteError: unable to open database file`
  (`SQLITE_CANTOPEN`) when `/data` was a bind mount to a root-owned host
  directory** (#57). The image chowned `/data` to `node` at build time, but a
  bind mount keeps its host ownership at runtime, so the `node` user couldn't
  create `users.db`. The container now starts as root, an entrypoint chowns the
  data dir (`CASUAL_LOCAL_PATH`, default `/data`) to `node`, then drops to `node`
  via `su-exec` before launching the server. Named volumes, fresh bind mounts,
  and `--user` overrides all still work; the server still runs unprivileged.

## [0.3.2] — 2026-06-20

Docker app refresh — rebuilds and republishes the image (so
`docker pull casualoffice/sheets:latest` works again, #57) and folds in
the editor + branding work that landed since 0.3.1.

### Added

- **Excel active-cell crosshair highlight** and a **zen (full-screen)
  cell editor**, wired from Univer OSS via context-menu triggers.

### Changed

- **Univer fork upgraded 0.24 → 0.25** under the hood.
- **Renamed to the CasualOffice org** — Docker image
  `casualoffice/sheets`, npm scope `@casualoffice/*`, in-app URLs, and the
  `sheet.casualoffice.org` subdomain.

### Fixed

- Republished the Docker image so the `latest` / `0.3` / `0` tags pull
  cleanly (#57).
- Gave the co-edit compaction e2e generous CI timeouts to stop flaky
  failures.

[GitHub release](https://github.com/CasualOffice/sheets/releases/tag/v0.3.2)

## [0.3.1] — 2026-06-11

Patch release — fixes the docker image so personal mode (Phase C)
actually boots.

### Fixed

- **`CASUAL_PERSONAL_MODE=single` / `multi` crashed the server on
  first boot** with `SqliteError: unable to open database file`. The
  named `casual-data` volume the docker-compose mounts at `/data` is
  owned by root, but the server drops to `USER node` before bootstrap
  — SQLite couldn't open `users.db` for write, the healthcheck never
  passed, every signup attempt 502'd. Dockerfile now provisions
  `/data` with `node:node` ownership before dropping privileges, so
  the volume is writable as soon as the container boots
  (`8d1de8d`).
- **Default `CASUAL_PERSONAL_MODE` reverted from `single` to `none`**
  in the docker-compose. The recent flip to `single` was meant to
  surface Phase C to fresh users on `docker compose up`, but it
  gated the anonymous-room coedit suite (every coedit-* e2e spec
  failed under single-mode) and broke the live demo's default flow.
  Personal mode is now opt-in via a one-line `.env` entry; README's
  "Recommended" section walks through it (`c167646`).

### Added — packages/sdk extraction (#56)

- **`@casualoffice/sheets@0.4.0`** publishes the xlsx import path
  as `@casualoffice/sheets/xlsx`. Drive (and any future host) can
  load `.xlsx` into a Univer `IWorkbookData` snapshot via
  `xlsxToWorkbookData(bytes)` instead of vendoring the parser. The
  apps/web pipeline is unchanged — just imports the shared mappers
  from the SDK instead of `./style-mapping`, `./constants`, etc.
- The 12 import-side + shared files (parse-impl, style-mapping,
  every `*-resource.ts`, pivot-passthrough) moved into
  `packages/sdk/src/xlsx/`; the worker runs from a sibling
  `parser.worker.js` via the same renderChunk pattern the docx
  editor SDK shipped in 1.0.1.
- Export-side stays in apps/web for now — Phase B of #56 covers
  the outline / charts / pivots / sparklines extension-point design
  before workbookDataToXlsx can move cleanly.

### Build infra

- CI's typecheck / e2e / e2e-prod / deploy-pages now build the SDK
  before consuming it (apps/web imports `@casualoffice/sheets/xlsx`
  whose types resolve through `packages/sdk/dist/`). Dockerfile
  does the same in the `build-web` stage. Without these, fresh
  builds fail TS2307 because `packages/sdk/dist` is empty
  (`4dc4dc4`, `01555d0`, `e6b4542`).

## [0.3.0] — 2026-06-08

Minor release rolling up two major feature batches plus the
Univer-fork perf revamp.

### Added — Phase C: personal mode (#49)

- **Per-user file storage** — server stores files under
  `<root>/users/<userID>/`, scoped through `local.PerUserStores`.
  Files stay open across sessions: the same user logs in on a
  different machine and finds their workbooks where they left them.
- **Auth foundation** — bcrypt + SQLite at `<root>/.casual/users.db`,
  HMAC-signed session cookies (30-day TTL), `__Host-`-prefixed under
  `SECURE_COOKIES=true`.
- **Auth routes** — `POST /auth/signup` / `/auth/login` / `/auth/logout`,
  `GET /auth/me`.
- **File CRUD over HTTP** — `POST /files`, `GET /files`, `GET /files/{id}`,
  `PUT /files/{id}/contents`, `PATCH /files/{id}`, `DELETE /files/{id}`.
- **Profile** — `displayName`, `timezone`, `locale`, `avatarUrl`, free-form
  `prefs`. Identity stays in SQLite; extended fields land in a
  `.profile.json` sidecar.
- **CLI** — `casual-docs reset-password / list-users / promote / demote`.
- **Admin** — `GET /admin/users`, `DELETE /admin/users/{id}` behind
  `RequireAdmin`. First signup auto-promotes.
- **UI** — `PersonalAuthGate` (login / signup modal), `UserMenu` pill +
  dropdown, `ProfileSettingsDialog`.

### Added — Phase D: WOPI host (#49)

- **WOPI client** in `apps/server/src/host/wopi/` + JWT verifier with
  JWKS cache. `docID = base64url(wopiSrc)` keeps the gateway stateless.
- **Embed redirect** at `GET /wopi/host`; access_token threaded through
  the WS preflight.
- **`WopiFileSource`** (TS) — front-end probe order is
  WOPI → Personal → Browser.
- **Lock / Unlock** + per-room `RefreshLock` ticker (10 min default) so
  long sessions don't lose the host-side lock idle-out.

### Changed — Univer fork (#51)

- Fork now lives at `vendor/univer-revamp/` as a submodule on
  `casual-sheets/0.24`. All 49 `@univerjs/*` packages resolve through
  `pnpm.overrides` to avoid the Service2-suffix DI-collision class.
- Five perf patches cherry-picked onto the fork (most impactful: stop
  re-walking the visible span in `setStylesCache`).

### Added — Mobile chrome + misc

- Soft cell-count cap on print export (#50) — was OOMing the tab on
  big workbooks.
- Right-click context menu picks up **Format Cells…** (#52).
- E2E hardening: timeouts + retries on `coedit-share`, `coedit-compaction`,
  `charts-p1`, and the long personal happy-path spec.

### Added — `@casualoffice/sheets` SDK (separate package)

- New `packages/sdk` shipping `@casualoffice/sheets@0.2.0`
  (signing + iframe postMessage protocol). Univer-Sheets React wrapper
  to follow.

## [0.2.1] — 2026-05-26

Patch release — closes the last two known gaps from the
production-readiness audit:

### Added — measurement

- **WS-side load harness** at `apps/server/scripts/wsloadtest.ts`
  (`pnpm --filter @sheet/server wsload`). Drives `@hocuspocus/
  provider` from Node — same handshake + sync protocol as real
  browsers — and measures connection setup, broadcast latency
  (sender push → peer observe), and dropped records.
  Configurable VUs / duration / write cadence.
- **Measured ceiling** in `docs/LOAD_TEST.md`: 1500 concurrent
  WS clients across 500 rooms, sustained 350 updates/s aggregate,
  p99 broadcast latency 3.2 ms, **zero dropped records**. The
  capacity model's "~500 active docs single-process latency knee"
  prediction was overcautious by ~10×; real binding constraint
  at that size is RAM, not broadcast CPU.

### Fixed — drawing-sync regression (cross-peer charts + images)

A drawing (image / chart) inserted by peer A did not appear on
peer B. Root cause: the bridge's `deepRewriteUnitId` only swaps
unitIds at object **keys** — but the `sheet.mutation.set-drawing-
apply` mutation carries its unitId at **position [0] of a json1
op path** (a positional array, not an object). The op replayed
on the joiner kept the OWNER's unitId, json1.type.apply walked
a path that doesn't exist locally, threw a bare "Error" with no
message, the classifier landed it as PERMANENT, and the drawing
silently failed to propagate.

Fix: new `rewriteJson1OpPathUnitId` helper in `bridge-helpers.ts`
walks the op (single JSONOp or JSONOpList) and substitutes the
leading unitId. Wired into `rewriteUnitId` for the drawing
mutation id.

Spec: `tests/e2e/coedit-drawings.spec.ts` — previously skipped
with a wrong-hypothesis comment about `registerDrawingData`;
now unskipped and passing.

### Capacity model corrections

- Removed the "~500 active docs latency knee" claim — measurement
  shows broadcast has 10× more headroom than the model assumed.
- Reordered the bottleneck list: **file descriptors hit first**
  (Linux default 1024), then RAM, then Redis, then CPU pegging
  (which wasn't approached even at 1500 concurrent WS).

### Test coverage

139 → 145 (+6 new tests for `rewriteJson1OpPathUnitId`).

## [0.2.0] — 2026-05-26

**The "production-readiness" release.** Six engineering streams that
turn v0.1's "real persistence + self-host story" into a workload you
can put real users on. Co-edit divergence becomes recoverable; the
gateway gets per-IP throttling + a hard room cap; we have measured
baseline numbers + a sizing model; and the FUniver-boundary type
debt starts coming down.

### Added — co-edit reliability

- **Bridge replay retry with backoff + dead-letter ring buffer**
  (`apps/web/src/collab/replay-retry.ts`). Replay failures are
  classified `transient` (dynamic-import chunk-load failures —
  retry with 300/900/2700 ms backoff) or `permanent` (malformed
  params / unknown command id — dead-letter immediately). Final
  failures append to a capped (20) ring buffer exposed via
  `BridgeHandle.getReplayDeadLetter()` / `subscribeReplayDeadLetter()`.
- **Click-to-expand replay-failure detail** in the `CollabIndicator`
  pill. Shows the last 5 dead-letter entries with mutation id,
  classification chip, truncated error, and age. Closes on
  outside-click / Escape. Auto-clears when the dead-letter empties.

### Added — backend hardening

- **Per-IP rate limit** via `@fastify/rate-limit`. New env vars:
  - `RATE_LIMIT_ENABLED` (default `true`) — master switch.
  - `RATE_LIMIT_PER_MIN` (default `60`) — applies to `POST /api/rooms`.
  - `UPLOAD_RATE_LIMIT_PER_MIN` (default `12`) — applies to
    `POST /api/rooms/:id/seed` and `POST /api/rooms/:id/snapshot`.
  Returns standard `429` + `retry-after` + `x-ratelimit-*` headers
  on overflow. Read endpoints (GET /snapshot) are NOT rate-limited.
- **Hard cap on concurrent rooms** via new `MAX_ROOMS` env
  (default `256`). When `create()` would exceed the cap, LRU-evicts
  the oldest **evictable** room (no password / no seed / no
  snapshot). If every slot is non-evictable, returns
  `503 capacity_full` + `retry-after: 60`. Two-pass eviction
  policy: prefer idle-but-evictable, fall back to live-but-evictable
  by `createdAt` — prevents a "spam open rooms" pattern from
  permanently locking out new users.
- **Boot log** of room registry + upload limits so operators can
  verify the configured caps at startup.

### Added — measurement + capacity planning

- **In-tree HTTP load harness** at `apps/server/scripts/loadtest.ts`
  (~190 lines, no new deps — uses Node's built-in `fetch` +
  `perf_hooks` + `FormData` + `Blob`). Drives the four bounded
  write-path endpoints with configurable VUs / duration / target;
  output is a grep-friendly numbers table. Run with
  `pnpm --filter @sheet/server load`.
- **v0.1 baseline numbers** documented in `docs/LOAD_TEST.md`:
  ~1900 req/s sustained, p99 < 3 ms across all four write endpoints
  with rate-limit disabled. Rate-limit verification run shows the
  bucket clamps a single IP exactly at the configured `60/min` +
  `12/min` envelopes.
- **Capacity model + sizing tiers** in `docs/CAPACITY_MODEL.md`.
  Workload-anchored: per-doc RAM / CPU / network / storage cost
  derived from the baseline + Yjs / Hocuspocus fan-out math. Five
  deployment tiers (Solo / Small / Mid / Big single-process /
  Sharded) with concrete dollar costs ($5/mo → $300/mo → linear).
  Worked example for a 4 vCPU / 8 GB / 180 SSD DigitalOcean
  General-Purpose droplet at 1 user/doc: 5 000–8 000 concurrent
  single-process, ~10 000–15 000 with cluster mode + sticky routing.
- **Production-pipeline doc** at `docs/PRODUCTION_PIPELINE.md` —
  rolling roadmap of the post-v0.1 reliability + hardening +
  measurement + release streams.

### Added — UX (toast + a11y + mobile + clarity)

- **Unified toast surface** (`apps/web/src/shell/toast/`) — `info`
  / `success` / `error` kinds, optional action button, accessible
  `role="status"` / `role="alert"`. Wired into:
  - File > Save / Export (success + error per format)
  - Autosave > Restore (success + error)
  - Insert Chart (`Added Chart 3`)
  - Sheet tab actions: rename ("Renamed to X"), duplicate, hide
    ("Hid X" with one-click `Show` action), delete ("Deleted X"
    with 8 s `Undo` action that calls Univer's command-stack undo).
  - Print Area set/clear (with `Undo` action that restores the
    previous range).
  - Paste Special apply (`Pasted: Formats` / `Column widths` /
    etc. — names the variant the user picked).
  - Flash Fill — outcome-aware (success carries the cell count;
    each failure mode gets a specific explanation rather than
    silently no-op'ing).
  - Save Version (success with `Open history` action; error catch).
  - Insert Sparkline (success names the type + anchor; error catch).
- **Peer count + queued-mutation count** in the `CollabIndicator`:
  - "Live · 2" when co-editing with 2 peers.
  - "Reconnecting · 3" when 3 of your edits are queued locally.
- **Humanised open-file errors** in the loading overlay — 8
  classifier branches (corrupt zip, encrypted, network, HTTP
  404 / 403 / 5xx, ods loader, memory) with the raw error
  collapsed under a `<details>`.
- **Insert Chart range error elevated to a banner** above the
  input with `role="alert"` + `aria-live` + `aria-invalid` on
  the input.
- **Ribbon group landmarks** — `role="group"` + `aria-label` on
  each ribbon group so screen readers announce boundaries.
- **Mobile fixes**: side-panel back-out pill is now unmistakable on
  touch (40 × 40 px "← Back"); toolbar overflow chevrons pinned to
  viewport edges so they don't get hidden behind the device
  notch; desktop toolbar hides correctly at ≤ 480 px.

### Changed — type-safety refactor (rolling)

- **Typed Univer facade** at `apps/web/src/univer-facade.ts`
  (~210 lines). Centralises the `as any` casts at the
  `FUniver → workbook → sheet → range` boundary into one
  auditable module. Surface: `sheetId`, `isHidden`, `maxRows`,
  `maxColumns`, `rangeAt`, `rangeBox`, `rangeFromA1`,
  `activateRange`, `dataRangeOrActive`, `setActiveSheet`,
  `findSheetById`, `saveWorkbook`, `activeSheet`, `activeRange`,
  `injector`, `viteEnv`, `viteEnvNumber`, `windowStringGlobal`.
- Converted 5 highest-traffic files (`tab-actions`, `sheet-actions`,
  `flash-fill`, `MenuBar`, `CollabDriver`) — **27 caller-side
  as-any sites eliminated**, 23 centralised in the facade. The
  remaining ~21 unconverted files are mechanical follow-up
  tracked under the rolling B1 stream.

### Fixed

- Formula bar didn't trigger initial recalc on workbook mount + swap
  (back-ported in v0.1.1; recorded here for completeness).
- Excel-style typed input (`$1,234` · `15%` · `(500)` · `€99`)
  parses as numbers instead of strings (v0.1.1 back-port).

### Internal

- 6 new unit tests for `replay-retry.ts` (classifier + retry
  scheduler + ring-buffer eviction).
- 6 new unit tests for `RoomRegistry` cap + LRU eviction.
- 10 new unit tests for toast normalisation + humanised errors
  (back-fill from v0.1.x pre-release).
- Total: **139 / 139 unit tests pass** (was 116 at start of cycle).

## [0.1.1] — 2026-05-25

Patch release — e2e stability + mobile + formula + Excel-style input.

### Fixed

- Excel-style typed input (`$1,234`, `15%`, `(500)`, `€99`) now
  parses as numbers instead of strings.
- Formula engine: initial recalc on workbook mount + swap so
  freshly-loaded docs show computed values immediately.
- Mobile: desktop toolbar correctly hides at ≤ 480 px viewport
  (regression of v0.0.6's Polish #4).
- Autosave restore banner was being clipped by the `.app` grid
  overflow — root cause of intermittent 1h+ e2e timeouts.
- E2E `waitForUniver` now dismisses the home screen on
  `loading-overlay-step` to unblock long CI runs.
- Various test alignment for the post-v0.1.0 mobile pass.

## [0.1.0] — 2026-05-24

The first version-bumped release. Real persistence (WOPI host:
memory / local / S3 / Postgres), JWT-secured access, runtime admin
panel, OCI image labels + rolling tags, complex pivot cache
passthrough (audit 54/54 pristine), full self-hosting +
customization docs section. Mobile lane back-ported.

Long-form notes:
[casualoffice.org/changelog/sheets-v0.1.0/](https://casualoffice.org/changelog/sheets-v0.1.0/) ·
[GitHub release](https://github.com/CasualOffice/sheets/releases/tag/v0.1.0)

## [0.0.6] — 2026-05-23 and earlier

See [`https://casualoffice.org/changelog/`](https://casualoffice.org/changelog/)
for per-release notes:
v0.0.6 / v0.0.5 / v0.0.4 / v0.0.3 / v0.0.2.
