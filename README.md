# Casual Sheets

[![CI](https://github.com/schnsrw/sheets/actions/workflows/ci.yml/badge.svg?branch=main)](https://github.com/schnsrw/sheets/actions/workflows/ci.yml)
[![Deploy](https://github.com/schnsrw/sheets/actions/workflows/deploy-pages.yml/badge.svg?branch=main)](https://github.com/schnsrw/sheets/actions/workflows/deploy-pages.yml)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](./LICENSE)
[![Node](https://img.shields.io/badge/node-%E2%89%A518.17-brightgreen)](#develop)
[![E2E tests](https://img.shields.io/badge/e2e-141%20passing-brightgreen)](./tests/e2e)
[![Docker](https://img.shields.io/badge/docker-schnsrw%2Fcasual--sheets-blue?logo=docker)](https://hub.docker.com/r/schnsrw/casual-sheets)

**Live demo: <https://sheet.schnsrw.live/>** — auto-deployed from `main` on every push.

A web-based, Excel-flavored spreadsheet editor with real-time co-editing. Built on [Univer](https://github.com/dream-num/univer) OSS (Apache-2.0).

The goal: feel like Excel, not like Google Sheets — ribbon, formula bar, file-centric workflow. Co-edit when you want it, single-user when you don't.

---

## Status

- **Phase 1** — single-user editor, feature-complete and locked down by **141 Playwright tests**.
- **Phase 2** — real-time co-editing (Yjs + Hocuspocus + Redis persistence), shipped as a self-hosted Docker image.
- **Phase 2.2** — polished share UX, live cursors with usernames, live-typing ghost overlay, large-file pipeline (worker-side xlsx parse + export, lazy-loaded plugins, halved in-memory footprint, loading overlay with phase-aware progress).

The hosted demo at <https://sheet.schnsrw.live/> ships single-user. Co-editing is gated by `VITE_COLLAB_ENABLED=1` and only present in the Docker image — open `/r/<roomId>` there to use it.

| Working | Coming |
| --- | --- |
| Office-style ribbon (Home / Insert / Formulas / Data / View / Review) | Charts |
| Inline cell editing, F2, Backspace, Delete, Escape | Pivot tables |
| Formula bar with editable Name Box (type `B5` to jump) | More fidelity on hyperlinks / pivots in xlsx round-trip |
| Fonts, colors, fill, wrap, alignment | Recent-files / landing page |
| Borders (split-button dropdown with 7 modes + color picker) | WOPI host integration |
| Cell merge + unmerge | Op-log compaction (Stage 6) for long-lived rooms |
| Group / outline rows + columns (collapse / expand) | |
| AutoSum / Average / Count / Min / Max | |
| Sort, Multi-column sort dialog, Filter | |
| Open / Save .xlsx / .ods / .csv / .tsv (full round-trip, workerized) | |
| Drag-and-drop file open + Save / Export toasts | |
| Hyperlinks (import + export round-trip, inline cell.p encoding) | |
| Drag-fill handle, Ctrl+D / Ctrl+R, relative-ref formula extension | |
| Tables (Format as Table, themes, Tables panel) | |
| Comments (with corner indicator markers) | |
| Freeze panes (top row / first column / at selection) | |
| Sheet tabs at the bottom (add / rename / delete / reorder) | |
| Print active sheet with Page Setup (orientation + margins, prefs persist) | |
| File menu with Properties dialog | |
| Office-style loading overlay for multi-MB opens | |
| Dynamic workbook growth — 1024×26 → 8192×1024 | |
| Help → Report a bug (prefills env to GitHub issue form) | |
| Material Symbols icons, Inter typography | |

### Co-editing

When the Docker image is running:

- **Share dialog** (File → Share for co-editing) — optional file seed, optional password, role choice (edit / view-only). Returns two copyable URLs.
- **Presence avatars** in the title bar — your initials + each peer's, with "Active now / Last seen Ns ago" tooltips. Idle peers fade.
- **Live cursors** on the grid — each peer's selected range, in their color, with name label, tracking scroll.
- **Live-typing ghost** — peers see characters appear in your cell as you type, not just on commit.
- **Password-protected rooms** with SHA-256 + constant-time compare; bad password closes the WS upgrade with 401.
- **Joiner fast-path**: on `/r/<id>` the joiner fetches a pre-parsed snapshot from the server (gzip-streamed, immutable-cached) and skips the xlsx parse entirely.

See [`docs/CO-EDITING.md`](./docs/CO-EDITING.md) for the architecture.

### Large-file pipeline

See [`docs/LARGE_FILE_PIPELINE.md`](./docs/LARGE_FILE_PIPELINE.md) for the staged plan; what's shipped today:

- **xlsx parse + export in Web Workers** — main thread never blocks on multi-MB opens or saves. ExcelJS lives only in the worker chunks; main bundle is 6.3 MB (was 8.3 MB).
- **Lazy plugin loading** — CF / DV / hyperlink / table / note / thread-comment / drawing / sort / filter / find-replace each ship as their own chunk, loaded after Univer mounts. Eager-loaded on snapshot inspection if the workbook needs them.
- **Snapshot-as-ref** — full `IWorkbookData` no longer lives in React state alongside Univer's copy. Cuts ~30–50% off peak heap on big files.
- **Hyperlinks inline in `cell.p`** — no more O(N) `AddHyperLinkCommand` replay per link at mount.
- **Op-log batching** — collab bridge coalesces mutations per microtask via `doc.transact`. One paste / sort / fill = one Yjs encode = one WS frame.
- **Selection-stats cap** — Cmd+A on a million-row sheet no longer freezes the UI computing Count / Sum / Avg over millions of cells.
- **Profiling harness** — `apps/web/src/perf.ts` + `perf-harness.spec.ts` wrap and verify the timed hot paths.
- **Loading overlay** — phase-aware ("Reading file…" → "Parsing…" → "Loading into editor…") with an elapsed timer and a "big file" hint after 4 s.

---

## Self-host with Docker

Single image. Web + Hocuspocus + Fastify in one container, Redis alongside for room persistence so sessions survive restarts.

```sh
# Run the published image directly (in-memory rooms — fine for a quick try):
docker run --rm -p 3000:3000 schnsrw/casual-sheets:latest
# open http://localhost:3000

# Or with persistence + Redis via compose:
docker compose up -d
# open http://localhost:3000
```

Co-edit a sheet:

1. Open `http://localhost:3000` in one tab. File → **Share for co-editing…** to set a password + role and get two share URLs.
2. Paste either URL into another browser / device. Owner sees joiners light up in the title bar avatar stack within ~1 s.
3. Move your selection — peers see your cursor + name follow it. Start typing — peers see characters appear in your cell live.

| Endpoint | What it does |
| --- | --- |
| `GET  /` | Serves the built web app |
| `GET  /r/:roomId` | Same SPA; bridges into the named Y.Doc |
| `POST /api/rooms` | Allocates a fresh room id; `{password?}` for protected rooms |
| `GET  /api/rooms/:id/info` | Pre-flight: `{needsPassword, hasSeed, hasSnapshot, clients}` |
| `POST /api/rooms/:id/seed` | Multipart xlsx upload — the room's starting workbook |
| `GET  /api/rooms/:id/seed` | Download the xlsx seed |
| `POST /api/rooms/:id/snapshot` | Gzipped JSON snapshot upload — joiner fast-path |
| `GET  /api/rooms/:id/snapshot` | Joiner fetches the pre-parsed snapshot (immutable-cached) |
| `GET  /api/rooms` | Diagnostic — live rooms + client counts |
| `GET  /health` | Liveness probe (returns `{ok, ts, rooms}`) |
| `WS   /yjs` | Hocuspocus Yjs sync; query: `?room=<id>&p=<password>` |

### Configuration

See [`.env.example`](./.env.example) for the full set. Copy to `.env`
and `docker compose up` picks it up automatically. Quick reference:

| Env var | Where | Default | Meaning |
| --- | --- | --- | --- |
| `PORT` | server | `3000` | HTTP + WS port |
| `HOST` | server | `0.0.0.0` | Bind address |
| `REDIS_URL` | server | _unset_ | If set, Y.Docs persist to Redis with a 7-day TTL |
| `ROOM_TTL_MIN` | server | `15` | Minutes a room stays in memory after the last client leaves |
| `MAX_UPLOAD_MB` | server | `100` | Cap on multipart + raw-binary uploads (xlsx seed + gzipped snapshot) |
| `VITE_COLLAB_ENABLED` | web build | `1` (image) / unset (Pages) | Ships co-edit code in the bundle |
| `VITE_MAX_OPEN_MB` | web build | `100` | Hard reject for File → Open / drag-drop; larger crashes the tab |
| `VITE_SOFT_WARN_MB` | web build | `25` | Threshold for the up-front "this is a large workbook" hint |

The `VITE_*` vars are baked in at build time. Pass them via
`--build-arg` on `docker build` or the `args:` block in
`docker-compose.yml` to customize the image.

---

## Develop

Prereqs: Node ≥ 18.17, pnpm 10+.

```sh
pnpm install               # one-time
pnpm dev:web               # http://127.0.0.1:5273
pnpm dev:server            # http://127.0.0.1:3000  (HTTP + Hocuspocus WS at /yjs)
pnpm test:e2e              # Playwright (auto-starts the web dev server)
pnpm test:e2e:ui           # Playwright UI mode
pnpm lint                  # eslint
pnpm format                # prettier --write
pnpm typecheck             # tsc across packages
```

Co-editing in dev needs both the web dev server (`:5273`) and the standalone server (`:3000`). Set `VITE_COLLAB_WS_URL=ws://127.0.0.1:3000/yjs` in `apps/web/.env.local` and visit `http://127.0.0.1:5273/r/<any-id>`.

---

## Repo layout

```
.
├── apps/
│   ├── web/                ← Vite + React frontend
│   │   ├── src/
│   │   │   ├── collab/     ← Yjs bridge, presence, share dialog
│   │   │   ├── shell/      ← title bar, ribbon, menu bar, dialogs
│   │   │   ├── univer/     ← Univer plugin registration + lazy loader
│   │   │   └── xlsx/       ← worker-side import/export (ExcelJS)
│   │   └── tests/
│   └── server/             ← Fastify + Hocuspocus (rooms, seed, snapshot)
├── tests/e2e/              ← Playwright e2e suite (141 specs across 35 files)
├── docs/
│   ├── ARCHITECTURE.md         ← system design
│   ├── CO-EDITING.md           ← op-log + presence design
│   ├── LARGE_FILE_PIPELINE.md  ← staged perf plan
│   └── RESEARCH.md             ← Univer technical brief
├── vendor/univer/          ← read-only Univer 0.22.1 source clone (gitignored)
├── Dockerfile              ← multi-stage build (deps → build-web → runtime)
├── docker-compose.yml      ← app + Redis
├── PLAN.md                 ← phased build plan
└── CLAUDE.md               ← project guardrails
```

`vendor/univer/` is a local clone of `dream-num/univer` for source-level study and is excluded from version control. To bootstrap on a fresh checkout:

```sh
git clone --depth 1 https://github.com/dream-num/univer.git vendor/univer
```

---

## Stack

| Concern | Pick |
| --- | --- |
| Editor + formula engine | Univer OSS (`@univerjs/core` + sheets plugins, pinned to 0.22.1) |
| Frontend | React + Vite + TypeScript (strict) |
| Lint / format | ESLint 9 (flat config) + Prettier |
| xlsx I/O | ExcelJS, run in dedicated Web Workers |
| ods / csv / tsv I/O | SheetJS Community (`@e965/xlsx`) |
| Icons / type | Material Symbols Outlined + Inter (Google Fonts) |
| E2E tests | Playwright (Chromium) |
| Collab transport | Yjs + Hocuspocus over WebSocket |
| Collab server | Fastify + raw `ws` (upgrade routing) |
| Collab persistence | Redis (optional, 7-day TTL on Y.Doc binary updates) |
| Formula offload | `UniverRPCMainThreadPlugin` ↔ `UniverRPCWorkerThreadPlugin` |

---

## What we explicitly don't do

Per [`CLAUDE.md`](./CLAUDE.md):

- **No persistence / accounts / WOPI** — anonymous sessions by room URL; in-memory + Redis only.
- **No AI / LLM** — the command bus is extensible; plug your own model in later if you want.
- **No mobile** — desktop browsers only (we ship a responsive shell down to 480 px, but the grid UX assumes a pointer).
- **No Univer Pro** — features missing from OSS (charts, pivots) get built here or deferred, not vendored.
- **No forking Univer** — `vendor/univer/` is read-only reference. Every fix sits on top.

---

## License

Apache-2.0. See [`LICENSE`](./LICENSE) and [`NOTICE`](./NOTICE) for attribution.

Vendored Univer source (`vendor/univer/`) keeps its upstream Apache-2.0 license; we do not modify it and it is not part of our build.
