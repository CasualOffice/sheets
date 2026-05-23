<div align="center">

# Casual Sheets

**Excel-flavored web spreadsheet with real-time collaborative editing**

[![CI](https://github.com/schnsrw/sheets/actions/workflows/ci.yml/badge.svg?branch=main)](https://github.com/schnsrw/sheets/actions/workflows/ci.yml)
[![Deploy](https://github.com/schnsrw/sheets/actions/workflows/deploy-pages.yml/badge.svg?branch=main)](https://github.com/schnsrw/sheets/actions/workflows/deploy-pages.yml)
[![Docker Pulls](https://img.shields.io/docker/pulls/schnsrw/casual-sheets?logo=docker)](https://hub.docker.com/r/schnsrw/casual-sheets)
[![Image Size](https://img.shields.io/docker/image-size/schnsrw/casual-sheets/latest?logo=docker&label=image)](https://hub.docker.com/r/schnsrw/casual-sheets)
[![E2E Tests](https://img.shields.io/badge/e2e-357%20passing-brightgreen?logo=playwright)](./tests/e2e)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](./LICENSE)

[**Live Demo →**](https://sheet.schnsrw.live/) &nbsp;·&nbsp; [Docker Hub →](https://hub.docker.com/r/schnsrw/casual-sheets) &nbsp;·&nbsp; [Architecture →](./docs/ARCHITECTURE.md)

</div>

---

Casual Sheets is a self-hostable, browser-based spreadsheet that looks and behaves like Microsoft Excel — ribbon UI, formula bar, file-centric workflow — with real-time multi-user co-editing built in. Upload an `.xlsx` file, share a link, and edit together instantly. No accounts, no database, no lock-in.

Built on [Univer OSS](https://github.com/dream-num/univer) (Apache-2.0) with a custom Office-style shell layered on top.

---

## ✨ What's Inside

### Spreadsheet Engine

- **Google-Docs-style title bar** — logo + filename + classic menus (File / Edit / View / Insert / Format / Data) in a single chrome strip, right-edge **panel rail** for Tables / Charts / Outline / Comments / History
- **Formula bar** with an editable Name Box (jump to any cell by typing its address)
- Cell editing: inline edit, F2, formula entry, multi-line paste, cross-sheet references with autocomplete
- Fonts, colors, fill, borders (7 modes + color picker), alignment, wrap, merge/unmerge
- Rows + columns: hide/unhide, group/outline (collapse/expand), resize
- **Freeze panes** — top row, first column, or freeze at any selection
- **Sort** (single and multi-column dialog), **Filter** (AutoFilter + re-apply), **Tables** (Format as Table, named tables panel)
- **Comments** with corner indicator markers
- **Conditional formatting**, **data validation**, **drawings** — fully round-tripped and co-edit synced
- **Charts** — 8 types, drag-to-resize, format dialog with **trendlines**, **date-axis detection**, **per-series colour overrides**; collab-synced; PNG embed in `.xlsx`
- **Pivot tables** — group-by + multi-aggregations, **filter fields**, **Refresh PivotTables**, **drill-down to source rows** (Ctrl+Shift+D), Insert dialog
- **Sparklines** — in-cell mini-charts (line / column / win-loss) via Insert → Sparkline, persist through the `__casual_sheets_sparklines__` resource
- **Analysis tools** — Name Manager (Ctrl+F3), Flash Fill (Ctrl+E), Goal Seek (iterative solver)
- **Show Formulas** (Ctrl+`) — non-destructive overlay paints formula source on every formula cell
- **Print Area** — A1 field in Page Setup + File-menu "Set / Clear Print Area"
- **Paste Special** (Ctrl+Alt+V) — All / Values / Formulas / Formats / Column widths / All except borders
- **Recent Files** — IndexedDB landing screen surfaces the last 10 workbooks on a blank `Untitled`
- **Version history** — auto-snapshots every ~10 min while dirty + "Save version…" manual entries; preview a snapshot then Restore (captures a "before restore" crumb for undo)
- **Dark theme** — title-bar sun/moon toggle, bridged to Univer's `ThemeService` so the canvas chrome flips too
- **Status-bar customisation** — right-click selection stats to toggle which appear (Average / Count / Sum / Min / Max / Numerical Count)
- Sheet tabs: add, rename, delete, reorder, color; tab strip refreshes live when peers act
- **Autosave** to IndexedDB with a restore banner if the tab closes unexpectedly
- Dynamic workbook growth — starts at 1024×26, expands to 8192×1024 on demand
- File → Properties dialog, Help → Report a Bug (GitHub issue prefill)
- **Inline SVG icons** (~155 components) — sharp at every size, no font-load delay

### File I/O

| Format | Open | Save / Export |
| --- | :---: | :---: |
| `.xlsx` | ✅ | ✅ |
| `.ods` | ✅ | ✅ |
| `.csv` / `.tsv` | ✅ | ✅ |

- Parsed and serialised entirely in **Web Workers** — the main thread never blocks on multi-MB files
- ODS round-trip: styles, dimensions, freeze, hyperlinks, comments, defined names
- xlsx round-trip: conditional formatting, data validation, drawings, tab colors, named ranges, hyperlinks (inline `cell.p` encoding), chart PNGs

### Excel Keyboard Shortcuts

40+ canonical shortcuts wired: Ctrl+1 (Format Cells), Shift+F8 (Add to Selection), Ctrl+Shift+L (AutoFilter), **Ctrl+Alt+L** (re-apply filter), Alt+= (AutoSum), F9 (recalculate), Ctrl+W (close), Alt+F1 (chart), **Ctrl+E** (Flash Fill), **Ctrl+F3** (Name Manager), **Ctrl+Alt+V** (Paste Special), **Ctrl+`** (Show Formulas), **Ctrl+[ / Ctrl+]** (precedents / dependents), **Ctrl+Shift+D** (pivot drill-down), number-format shortcuts, hide/unhide rows/columns, border combos.

### Co-editing

Available in the Docker image. Single-user on the hosted demo.

- **Share dialog** — File → Share for co-editing. Set a password, choose edit or view-only, get two copyable URLs
- **Presence avatars** — title-bar avatar stack with "Active now / Last seen Ns ago" tooltips; idle peers fade
- **Live cursors** — each peer's selection range in their color, with a name label, tracking scroll and frozen panes
- **Live-typing ghost** — characters appear in the peer's cell as they type, not just on commit
- **Full mutation sync** — cell values, styles, structure, conditional formatting, data validation, drawings, workbook metadata (tab colors, zoom, freeze, sheet visibility) all propagate cross-peer
- **View-only enforcement** at the Univer engine layer — view-only joiners cannot mutate the workbook, not just at the UI level
- **Session-history panel** — per-room op log with timestamps; review or revert any change
- **Divergence detection** — amber "Out of sync" pill when state vectors diverge; "Waiting to reconnect" banner on WebSocket drop
- **Password-protected rooms** — SHA-256 + constant-time compare; wrong password → HTTP 401 on the WS upgrade
- **Op-log compaction** — long-lived rooms compact the Yjs update log on `requestIdleCallback` to keep memory bounded
- **Joiner fast-path** — gzip-streamed pre-parsed snapshot from the server; joiners skip the xlsx parse entirely

See [`docs/CO-EDITING.md`](./docs/CO-EDITING.md) for the full architecture.

---

## 🐳 Self-Host with Docker

A single multi-arch image (`linux/amd64` + `linux/arm64`). Web, Hocuspocus, and Fastify run in one container; Redis runs alongside for room persistence.

### Quick start (in-memory, no persistence)

```sh
docker run --rm -p 3000:3000 schnsrw/casual-sheets:latest
```

Open `http://localhost:3000`.

### Recommended: with Redis persistence

Paste this `docker-compose.yml` and run `docker compose up -d`:

```yaml
services:
  app:
    image: schnsrw/casual-sheets:latest
    restart: unless-stopped
    ports: ['3000:3000']
    environment:
      REDIS_URL: redis://redis:6379
      ROOM_TTL_MIN: '15'
    depends_on:
      redis:
        condition: service_healthy

  redis:
    image: redis:7.4-alpine
    restart: unless-stopped
    command: ['redis-server', '--appendonly', 'yes']
    volumes:
      - redis-data:/data
    healthcheck:
      test: ['CMD', 'redis-cli', 'ping']
      interval: 10s
      timeout: 3s
      retries: 5

volumes:
  redis-data:
```

### Try co-editing

1. Open `http://localhost:3000`. **File → Share for co-editing…** to set a password and get two share URLs.
2. Paste either URL into another browser or device — the joiner connects in under a second.
3. Move your selection — peers see your named cursor follow it. Type in a cell — peers see characters appear live.

### API surface

| Method | Path | Description |
| --- | --- | --- |
| `GET` | `/` | Serves the built web app |
| `GET` | `/r/:roomId` | Same SPA; bridges into the named Y.Doc |
| `POST` | `/api/rooms` | Create a room; body `{password?}` |
| `GET` | `/api/rooms/:id/info` | Pre-flight: `{needsPassword, hasSeed, hasSnapshot, clients}` |
| `POST` | `/api/rooms/:id/seed` | Multipart `.xlsx` upload — the room's starting workbook |
| `GET` | `/api/rooms/:id/seed` | Download the seed file |
| `POST` | `/api/rooms/:id/snapshot` | Gzipped JSON snapshot — joiner fast-path upload |
| `GET` | `/api/rooms/:id/snapshot` | Joiner snapshot fetch (immutable-cached) |
| `GET` | `/api/rooms` | Diagnostic: live rooms + client counts |
| `GET` | `/health` | Liveness probe — `{ok, ts, rooms}` |
| `WS` | `/yjs` | Hocuspocus sync; `?room=<id>&p=<password>` |

### Configuration

Copy `.env.example` to `.env`; `docker compose up` picks it up automatically.

| Env var | Scope | Default | Description |
| --- | --- | --- | --- |
| `PORT` | server | `3000` | HTTP + WebSocket listen port |
| `HOST` | server | `0.0.0.0` | Bind address |
| `REDIS_URL` | server | _unset_ | Redis connection string; enables Y.Doc persistence with a 7-day TTL |
| `ROOM_TTL_MIN` | server | `15` | Minutes a room stays alive after the last client leaves |
| `MAX_UPLOAD_MB` | server | `100` | Upload cap for xlsx seed + gzipped snapshot |
| `VITE_COLLAB_ENABLED` | build | `1` in image | Include co-edit code in the bundle |
| `VITE_MAX_OPEN_MB` | build | `100` | Hard reject threshold for File → Open / drag-drop |
| `VITE_SOFT_WARN_MB` | build | `25` | Shows a "large file" hint before opening |

`VITE_*` vars are baked in at build time. Pass them with `--build-arg` on `docker build`, or via the `args:` block in `docker-compose.yml`.

---

## 🛠 Develop

**Prerequisites:** Node ≥ 18.17, pnpm 10+

```sh
pnpm install               # install workspace dependencies
pnpm dev:web               # Vite dev server  →  http://127.0.0.1:5273
pnpm dev:server            # Fastify + Hocuspocus  →  http://127.0.0.1:3000
pnpm test:e2e              # run Playwright suite (auto-starts web dev server)
pnpm test:e2e:ui           # Playwright UI mode
pnpm lint                  # ESLint
pnpm format                # Prettier --write
pnpm typecheck             # tsc across all packages
```

**Co-editing in dev** requires both servers running. Add to `apps/web/.env.local`:

```env
VITE_COLLAB_WS_URL=ws://127.0.0.1:3000/yjs
```

Then open `http://127.0.0.1:5273/r/<any-room-id>` in two tabs.

---

## 📁 Repo Layout

```
.
├── apps/
│   ├── web/                     # Vite + React frontend
│   │   └── src/
│   │       ├── collab/          # Yjs bridge, presence, share dialog
│   │       ├── shell/           # title bar, ribbon, menu bar, dialogs
│   │       ├── univer/          # Univer plugin registration + lazy loader
│   │       └── xlsx/            # worker-side import/export (ExcelJS)
│   └── server/                  # Fastify + Hocuspocus (rooms, seed, snapshot)
├── tests/e2e/                   # Playwright e2e suite — 357 tests across 87 files
├── docs/
│   ├── ARCHITECTURE.md          # system design
│   ├── CO-EDITING.md            # op-log + presence design
│   ├── LARGE_FILE_PIPELINE.md   # staged perf plan
│   └── RESEARCH.md              # Univer technical brief
├── vendor/univer/               # our univer-revamp fork (v0.24.0, gitignored, full clone for source-level edits)
├── Dockerfile                   # multi-stage build (deps → build-web → runtime)
├── docker-compose.yml           # app + Redis
├── PLAN.md                      # phased build plan
└── CLAUDE.md                    # project guardrails for AI-assisted development
```

Bootstrap the fork on a fresh checkout (full clone — depth-1 won't carry the history needed for fork commits):

```sh
git clone git@github.com:schnsrw/univer-revamp.git vendor/univer
```

---

## 🧱 Stack

| Concern | Choice |
| --- | --- |
| Grid + formula engine | Univer OSS (`@univerjs/core` + sheets plugins, pinned to 0.22.1) |
| Frontend | React 18 + Vite + TypeScript (strict mode) |
| Styling | Tailwind CSS + Material Symbols Outlined + Inter (Google Fonts) |
| Lint / format | ESLint 9 flat config + Prettier |
| xlsx I/O | ExcelJS in dedicated Web Workers |
| ods / csv / tsv I/O | SheetJS Community (`@e965/xlsx`) |
| E2E tests | Playwright (Chromium) |
| Collab transport | Yjs (CRDT) + Hocuspocus over WebSocket |
| Collab server | Fastify + raw `ws` (direct WS upgrade routing) |
| Persistence | Redis — optional, 7-day TTL on Y.Doc binary updates |
| Formula offload | `UniverRPCMainThreadPlugin` ↔ `UniverRPCWorkerThreadPlugin` |

---

## 🚫 Explicit Non-Goals

- **No persistence / accounts / WOPI** — anonymous sessions by room URL; in-memory + Redis only.
- **No AI / LLM features** — the Univer command bus is extensible; wire your own model in later.
- **No mobile** — desktop browsers only. The shell is responsive to 480 px, but the grid UX assumes a pointer device.
- **No Univer Pro** — everything is built on OSS. Missing features are built here or deferred; the commercial Pro package is never used.

---

## 📄 License

Apache-2.0. See [`LICENSE`](./LICENSE) and [`NOTICE`](./NOTICE).

Our fork of Univer (`vendor/univer/` → [`schnsrw/univer-revamp`](https://github.com/schnsrw/univer-revamp)) retains its upstream Apache-2.0 license. It is gitignored from this repo and is not part of this project's build today; modifications land in the fork repo independently. See [`docs/UNIVER_FORK_PERF.md`](./docs/UNIVER_FORK_PERF.md) for the active perf-improvement plan.
