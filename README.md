<div align="center">

# Casual Sheets

**Open-source self-hosted web spreadsheet with `.xlsx` round-trip and real-time co-editing — an alternative to Google Sheets, Excel Online, and OnlyOffice you run on your own server.**

[![CI](https://github.com/CasualOffice/sheets/actions/workflows/ci.yml/badge.svg?branch=main)](https://github.com/CasualOffice/sheets/actions/workflows/ci.yml)
[![Deploy](https://github.com/CasualOffice/sheets/actions/workflows/deploy-pages.yml/badge.svg?branch=main)](https://github.com/CasualOffice/sheets/actions/workflows/deploy-pages.yml)
[![Docker Pulls](https://img.shields.io/docker/pulls/casualoffice/sheets?logo=docker)](https://hub.docker.com/r/casualoffice/sheets)
[![Image Size](https://img.shields.io/docker/image-size/casualoffice/sheets/latest?logo=docker&label=image)](https://hub.docker.com/r/casualoffice/sheets)
[![E2E Tests](https://img.shields.io/badge/e2e-398%20passing-brightgreen?logo=playwright)](./tests/e2e)
[![Unit Tests](https://img.shields.io/badge/unit-235%20passing-brightgreen)](./apps)
[![Version](https://img.shields.io/badge/version-v0.3.2-blue)](./CHANGELOG.md)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](./LICENSE)

[**Live Demo →**](https://sheet.casualoffice.org/) &nbsp;·&nbsp; [Docker Hub →](https://hub.docker.com/r/casualoffice/sheets) &nbsp;·&nbsp; [Architecture →](./docs/ARCHITECTURE.md) &nbsp;·&nbsp; [Embed / SDK →](./docs/SDK_ARCHITECTURE.md) &nbsp;·&nbsp; [Comparisons →](https://casualoffice.org/vs/)

</div>

---

Casual Sheets is a **self-hostable, browser-based spreadsheet** that looks and behaves like Microsoft Excel — ribbon UI, formula bar, file-centric workflow — with **real-time multi-user co-editing** built in. Upload an `.xlsx` file, share a link, and edit together instantly. **No accounts, no Microsoft / Google login, no lock-in.** One Docker container, runs on a $5/mo VPS, scales to ~5 000 concurrent users on a $48/mo box (numbers [measured](./docs/LOAD_TEST.md), not hand-waved).

**Compares to:**
[Google Sheets](https://casualoffice.org/vs/sheets-vs-google-sheets/) ·
[Excel Online](https://casualoffice.org/vs/sheets-vs-excel-online/) ·
[OnlyOffice](https://casualoffice.org/vs/sheets-vs-onlyoffice/)

Built on [Univer OSS](https://github.com/dream-num/univer) (Apache-2.0) — the OSS variant, **never the Pro package** — with a custom Office-style shell layered on top. Sister projects: [Casual Editor](https://github.com/CasualOffice/docs) (`.docx`) and [Casual Slides](https://github.com/CasualOffice/slides) (`.pptx`).

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

| Format          | Open | Save / Export |
| --------------- | :--: | :-----------: |
| `.xlsx`         |  ✅  |      ✅       |
| `.ods`          |  ✅  |      ✅       |
| `.csv` / `.tsv` |  ✅  |      ✅       |

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
docker run --rm -p 3000:3000 casualoffice/sheets:latest
```

Open `http://localhost:3000`. **Anonymous mode** — no signup, no persistence; close the container and everything's gone. Good for kicking the tyres.

### Recommended: personal mode + Redis persistence

Paste this `docker-compose.yml` and run `docker compose up -d`:

```yaml
services:
  app:
    image: casualoffice/sheets:latest
    restart: unless-stopped
    ports: ['3000:3000']
    volumes:
      - casual-data:/data
    environment:
      # Phase C — sign up, save, reopen across restarts.
      CASUAL_STORAGE: local
      CASUAL_LOCAL_PATH: /data
      CASUAL_PERSONAL_MODE: single
      # Y.Doc snapshots for co-edit rooms.
      REDIS_URL: redis://redis:6379
      ROOM_TTL_MIN: '60'
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
  casual-data:
  redis-data:
```

Open `http://localhost:3000`, sign up — your account + workbooks live in the `casual-data` volume and survive `docker compose down`. `CASUAL_PERSONAL_MODE=multi` keeps signup open if you want more than one account on the same instance; `none` reverts to the anonymous shape.

For an admin without manual signup, set `CASUAL_BOOTSTRAP_USER=admin:hunter2` once in your `.env` — silently no-ops after the users table has a row. Full reference: [`docs/ENV.md`](./docs/ENV.md), [`docs/self-hosting/personal-mode.md`](./docs/self-hosting/personal-mode.md).

### Try co-editing

1. Open `http://localhost:3000`. **File → Share for co-editing…** to set a password and get two share URLs.
2. Paste either URL into another browser or device — the joiner connects in under a second.
3. Move your selection — peers see your named cursor follow it. Type in a cell — peers see characters appear live.

### API surface

| Method | Path                      | Description                                                  |
| ------ | ------------------------- | ------------------------------------------------------------ |
| `GET`  | `/`                       | Serves the built web app                                     |
| `GET`  | `/r/:roomId`              | Same SPA; bridges into the named Y.Doc                       |
| `POST` | `/api/rooms`              | Create a room; body `{password?}`                            |
| `GET`  | `/api/rooms/:id/info`     | Pre-flight: `{needsPassword, hasSeed, hasSnapshot, clients}` |
| `POST` | `/api/rooms/:id/seed`     | Multipart `.xlsx` upload — the room's starting workbook      |
| `GET`  | `/api/rooms/:id/seed`     | Download the seed file                                       |
| `POST` | `/api/rooms/:id/snapshot` | Gzipped JSON snapshot — joiner fast-path upload              |
| `GET`  | `/api/rooms/:id/snapshot` | Joiner snapshot fetch (immutable-cached)                     |
| `GET`  | `/api/rooms`              | Diagnostic: live rooms + client counts                       |
| `GET`  | `/health`                 | Liveness probe — `{ok, ts, rooms}`                           |
| `WS`   | `/yjs`                    | Hocuspocus sync; `?room=<id>&p=<password>`                   |

### Configuration

Copy `.env.example` to `.env`; `docker compose up` picks it up automatically.

| Env var               | Scope  | Default      | Description                                                         |
| --------------------- | ------ | ------------ | ------------------------------------------------------------------- |
| `PORT`                | server | `3000`       | HTTP + WebSocket listen port                                        |
| `HOST`                | server | `0.0.0.0`    | Bind address                                                        |
| `REDIS_URL`           | server | _unset_      | Redis connection string; enables Y.Doc persistence with a 7-day TTL |
| `ROOM_TTL_MIN`        | server | `15`         | Minutes a room stays alive after the last client leaves             |
| `MAX_UPLOAD_MB`       | server | `100`        | Upload cap for xlsx seed + gzipped snapshot                         |
| `VITE_COLLAB_ENABLED` | build  | `1` in image | Include co-edit code in the bundle                                  |
| `VITE_MAX_OPEN_MB`    | build  | `100`        | Hard reject threshold for File → Open / drag-drop                   |
| `VITE_SOFT_WARN_MB`   | build  | `25`         | Shows a "large file" hint before opening                            |

`VITE_*` vars are baked in at build time. Pass them with `--build-arg` on `docker build`, or via the `args:` block in `docker-compose.yml`.

The full env-var matrix (storage backends, networking, admin, JWT, webhooks, more) is documented in [`docs/ENV.md`](./docs/ENV.md) — that's the canonical reference the admin panel reads from too.

### 🧩 Embed the editor (SDK)

The editor is becoming an npm package — `@casualoffice/sheets` — that you mount as
a single React component. You own storage (the `onChange` snapshot stream), and
collaboration is opt-in; no backend is required to embed.

> **In active development, not yet published.** `@casualoffice/sheets` is the new
> Excalidraw-model editor SDK, landing on `main` now. The older
> `@schnsrw/casual-sheets@0.8.0` is published but is a pre-restructure line
> without this API. First `@casualoffice/sheets` publish ships the surface below.

```tsx
import { CasualSheets } from '@casualoffice/sheets/sheets';
import '@casualoffice/sheets/styles';

<CasualSheets initialData={snapshot} onChange={(snap) => persist(snap)} />;
```

Full guide — install, props, `CasualSheetsAPI`, xlsx import, opt-in collab:
[`docs/INTEGRATION.md`](./docs/INTEGRATION.md). Sandboxed `<iframe>` / signed
embeds: [`docs/SDK_SIGNING_EMBED.md`](./docs/SDK_SIGNING_EMBED.md).

### 📚 Self-hosting + customization docs

| Topic                                              | Lives on the site at                                                                                         | Source in this repo                                                          |
| -------------------------------------------------- | ------------------------------------------------------------------------------------------------------------ | ---------------------------------------------------------------------------- |
| **Self-hosting overview**                          | [casualoffice.org/docs/sheets/self-hosting/](https://casualoffice.org/docs/sheets/self-hosting/)             | [`docs/self-hosting/overview.md`](./docs/self-hosting/overview.md)           |
| Reverse-proxy recipes (nginx · Caddy · Traefik)    | [/docs/sheets/self-hosting-reverse-proxy/](https://casualoffice.org/docs/sheets/self-hosting-reverse-proxy/) | [`docs/self-hosting/reverse-proxy.md`](./docs/self-hosting/reverse-proxy.md) |
| TLS + custom domain                                | [/docs/sheets/self-hosting-tls/](https://casualoffice.org/docs/sheets/self-hosting-tls/)                     | [`docs/self-hosting/tls.md`](./docs/self-hosting/tls.md)                     |
| CORS                                               | [/docs/sheets/self-hosting-cors/](https://casualoffice.org/docs/sheets/self-hosting-cors/)                   | [`docs/self-hosting/cors.md`](./docs/self-hosting/cors.md)                   |
| Scaling                                            | [/docs/sheets/self-hosting-scaling/](https://casualoffice.org/docs/sheets/self-hosting-scaling/)             | [`docs/self-hosting/scaling.md`](./docs/self-hosting/scaling.md)             |
| Backups                                            | [/docs/sheets/self-hosting-backups/](https://casualoffice.org/docs/sheets/self-hosting-backups/)             | [`docs/self-hosting/backups.md`](./docs/self-hosting/backups.md)             |
| **Customization overview**                         | [/docs/sheets/customization/](https://casualoffice.org/docs/sheets/customization/)                           | [`docs/customization/overview.md`](./docs/customization/overview.md)         |
| Auth — JWT, roles, permissions, features           | [/docs/sheets/customization-auth/](https://casualoffice.org/docs/sheets/customization-auth/)                 | [`docs/customization/auth.md`](./docs/customization/auth.md)                 |
| Webhooks — events, payload, signature verification | [/docs/sheets/customization-webhooks/](https://casualoffice.org/docs/sheets/customization-webhooks/)         | [`docs/customization/webhooks.md`](./docs/customization/webhooks.md)         |

The admin panel at **`/admin`** is the runtime UI for everything above — branding, storage backend selection, networking, room limits, auth providers (JWT live; OIDC + SAML stubbed for v0.2), webhook subscriptions, base path. Enable it by setting `CASUAL_ADMIN_USERNAME` + `CASUAL_ADMIN_PASSWORD` + `CASUAL_JWT_SECRET` on the container.

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
├── vendor/univer-revamp/        # our Univer fork (v0.25.0 submodule, wired via pnpm.overrides)
├── Dockerfile                   # multi-stage build (deps → build-web → runtime)
├── docker-compose.yml           # app + Redis
├── PLAN.md                      # phased build plan
└── CLAUDE.md                    # project guardrails for AI-assisted development
```

Bootstrap the fork on a fresh checkout (full clone — depth-1 won't carry the history needed for fork commits):

```sh
git clone git@github.com:CasualOffice/univer-revamp.git vendor/univer
```

---

## 🧱 Stack

| Concern               | Choice                                                           |
| --------------------- | ---------------------------------------------------------------- |
| Grid + formula engine | Univer OSS (`@univerjs/core` + sheets plugins, pinned to 0.25.x) |
| Frontend              | React 18 + Vite + TypeScript (strict mode)                       |
| Styling               | Tailwind CSS + Material Symbols Outlined + Inter (Google Fonts)  |
| Lint / format         | ESLint 9 flat config + Prettier                                  |
| xlsx I/O              | ExcelJS in dedicated Web Workers                                 |
| ods / csv / tsv I/O   | SheetJS Community (`@e965/xlsx`)                                 |
| E2E tests             | Playwright (Chromium)                                            |
| Collab transport      | Yjs (CRDT) + Hocuspocus over WebSocket                           |
| Collab server         | Fastify + raw `ws` (direct WS upgrade routing)                   |
| Persistence           | Redis — optional, 7-day TTL on Y.Doc binary updates              |
| Formula offload       | `UniverRPCMainThreadPlugin` ↔ `UniverRPCWorkerThreadPlugin`      |

---

## 🚫 Explicit Non-Goals

- **No AI / LLM features** — the Univer command bus is extensible; wire your own model in later. We're an editor, not an AI product.
- **No Univer Pro** — everything is built on OSS. Missing features are built here or deferred; the commercial Pro package is never used.
- **No native mobile apps** — desktop browsers are first-class; mobile back-port (viewer + light editor) ships at `≤ 480 px` but isn't a native iOS / Android app.
- **No 1:1 Google-Sheets feature parity** — we don't chase every Power Query / LAMBDA / niche function. The bar is "the surface a typical business spreadsheet uses, round-tripped honestly."

---

## 📄 License

Apache-2.0. See [`LICENSE`](./LICENSE) and [`NOTICE`](./NOTICE).

Our fork of Univer (`vendor/univer/` → [`CasualOffice/univer-revamp`](https://github.com/CasualOffice/univer-revamp)) retains its upstream Apache-2.0 license. It is gitignored from this repo and is not part of this project's build today; modifications land in the fork repo independently. See [`docs/UNIVER_FORK_PERF.md`](./docs/UNIVER_FORK_PERF.md) for the active perf-improvement plan.
