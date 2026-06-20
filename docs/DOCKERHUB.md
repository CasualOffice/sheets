# Casual Sheets

**Excel-flavored web spreadsheet with real-time collaborative editing.**

Built on [Univer OSS](https://github.com/dream-num/univer) (Apache-2.0).
Single image — web app, Hocuspocus, and Fastify all in one container.

**Source:** [github.com/CasualOffice/sheets](https://github.com/CasualOffice/sheets) &nbsp;·&nbsp; **Demo (single-user):** [sheet.casualoffice.org](https://sheet.casualoffice.org/) &nbsp;·&nbsp; **Docs:** [casualoffice.org/docs/sheets/](https://casualoffice.org/docs/sheets/)

---

## Quick Start

```sh
# In-memory rooms — great for a quick try (rooms vanish on restart):
docker run --rm -p 3000:3000 casualoffice/sheets:latest
# open http://localhost:3000
```

### With Redis persistence (recommended)

```sh
docker compose up -d
# open http://localhost:3000
```

Paste this `docker-compose.yml` if you don't want to clone the repo:

```yaml
services:
  app:
    image: casualoffice/sheets:0.3 # rolling minor — auto picks up patch updates
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

---

## Co-Edit a Sheet

1. Open `http://localhost:3000`. **File → Share for co-editing…** — set a password + role, get two share URLs.
2. Paste either URL into another browser or device. Joiners appear in the title-bar avatar stack within ~1 s.
3. Move your selection — peers see your named cursor follow it. Type in a cell — peers see characters appear live.

Anonymous sessions — no accounts required. Rooms are addressed at `/r/<roomId>`.

---

## Features

- **Office-style ribbon** — Home, Insert, Formulas, Data, View, Review tabs
- **Full xlsx / ods / csv / tsv round-trip** — parsed and serialised in Web Workers; main thread never blocks
- **`.xlsm` macros + pivot tables round-trip byte-equal** via OOXML passthrough (audit 54 / 54 probes pristine)
- **Charts** — insert dialog, 8 chart types, trendlines, date-axis detection, drag/resize, format dialog, co-edit sync, PNG embed in `.xlsx`
- **Pivot tables** — multi-row compact layout, filter fields, drill-down (`Ctrl+Shift+D`)
- **Sparklines** — line / column / win-loss, in-cell, xlsx round-trip
- **Conditional formatting, data validation, drawings** — round-tripped and co-edit synced
- **Autosave** to IndexedDB with a restore banner on reload
- **Version history** — auto-snapshots with one-click restore (single-user mode)
- **Co-editing** — live cursors, live-typing ghost, presence avatars, password-protected rooms
- **View-only enforcement** at the Univer engine layer (server-side `onAuthenticate` gates)
- **Session-history panel** — per-room op log; review or revert changes
- **Divergence detection** — "Out of sync" pill + "Waiting to reconnect" banner
- **11-template home gallery** — Personal / Work / Finance / Education
- **Mobile viewer + light editor** — touch-pan canvas scroll, bottom action bar, compact chrome at ≤ 720 px / ≤ 480 px
- **357+ Playwright e2e tests** covering all major workflows

---

## Configuration

Quick reference; the canonical doc lives at [`docs/ENV.md`](https://github.com/CasualOffice/sheets/blob/main/docs/ENV.md).

| Env var         | Default   | Description                                                         |
| --------------- | --------- | ------------------------------------------------------------------- |
| `PORT`          | `3000`    | HTTP + WebSocket listen port                                        |
| `HOST`          | `0.0.0.0` | Bind address                                                        |
| `REDIS_URL`     | _unset_   | Redis connection string; enables Y.Doc persistence with a 7-day TTL |
| `ROOM_TTL_MIN`  | `15`      | Minutes a room stays alive after the last client disconnects        |
| `MAX_UPLOAD_MB` | `100`     | Upload cap for xlsx seed and gzipped snapshot                       |

Build-time knobs (`VITE_MAX_OPEN_MB`, `VITE_SOFT_WARN_MB`, `VITE_COLLAB_ENABLED`) are baked into the image. Override with `--build-arg` when building your own image.

WOPI host-integration env vars (`CASUAL_STORAGE`, `CASUAL_S3_*`, `CASUAL_PG_URL`, …) plus the admin / networking surfaces (`CASUAL_ADMIN_PASSWORD`, `CASUAL_PUBLIC_ORIGIN`, `CASUAL_CORS_ORIGINS`, `CASUAL_TRUST_PROXY`, …) are supported. See `docs/ENV.md` for the full table.

---

## Tags

The CI builds and publishes the **full rolling-tag set** on every release. Pick the cadence you want:

| Tag pattern | What you get       | Use when                                  |
| ----------- | ------------------ | ----------------------------------------- |
| `latest`    | The newest release | Local dev / "I want the bleeding edge"    |
| `0`         | Latest 0.x.y       | Reserved — bumps once v1.0.0 ships        |
| `0.3`       | Latest 0.3.x       | Patch updates only — recommended for prod |
| `0.3.3`     | Pinned exact       | Tightest production pin                   |

Multi-arch manifest: `linux/amd64` + `linux/arm64`. SBOM + provenance attestations ride along in the OCI manifest for `trivy` / `snyk` / GitHub dep-graph consumers.

### Release notes

Per-release changelog (newest first) lives in the repo:
[**CHANGELOG.md**](https://github.com/CasualOffice/sheets/blob/main/CHANGELOG.md) ·
[GitHub Releases](https://github.com/CasualOffice/sheets/releases). Latest line is
**0.3.x** (single-image web + Hocuspocus + Fastify, WOPI / personal / collab modes).

---

## OCI labels — `docker inspect` your build

Every published image carries `org.opencontainers.image.*` labels:

```sh
docker inspect casualoffice/sheets:latest \
  | jq '.[0].Config.Labels | with_entries(select(.key | startswith("org.opencontainers")))'
```

Sample output:

```json
{
  "org.opencontainers.image.title": "Casual Sheets",
  "org.opencontainers.image.description": "Excel-flavored web spreadsheet …",
  "org.opencontainers.image.url": "https://sheet.casualoffice.org/",
  "org.opencontainers.image.source": "https://github.com/CasualOffice/sheets",
  "org.opencontainers.image.documentation": "https://casualoffice.org/docs/sheets/",
  "org.opencontainers.image.vendor": "Sachin Sarwa",
  "org.opencontainers.image.licenses": "Apache-2.0",
  "org.opencontainers.image.version": "v0.3.3",
  "org.opencontainers.image.revision": "abc1234…",
  "org.opencontainers.image.created": "2026-06-20T14:23:09Z"
}
```

The `version`, `revision`, and `created` labels are baked in at CI build-time and travel with the image — pin a specific build by its full SHA + tag.

---

## What's Inside the Image

- **Web app** — React + Vite + TypeScript, served statically by the same Fastify server.
- **Collab server** — Yjs + Hocuspocus over WebSocket at `/yjs`. Per-mutation op-log with `fromCollab` echo-loop prevention and Stage-6 compaction.
- **xlsx I/O** — ExcelJS + JSZip running in dedicated Web Workers; main thread never blocks on multi-MB opens or saves.
- **OOXML passthrough** — `xl/vbaProject.bin` (macros) and `xl/pivotCaches/**` + `xl/pivotTables/**` (pivots) ride the round-trip byte-equal so `.xlsm` files keep macros and pivot files re-render as pivots in Excel.
- **Lazy plugin loading** — conditional formatting, data validation, hyperlinks, drawings, sort, filter, find-replace each load as their own chunk; eager-loaded on snapshot inspection.

---

## License

Apache-2.0. Vendored Univer source retains its upstream Apache-2.0 license; it is not modified and is not part of the build.
