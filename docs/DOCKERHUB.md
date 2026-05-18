# Casual Sheets

**A web-based, Excel-flavored spreadsheet with real-time co-editing.**
Built on [Univer OSS](https://github.com/dream-num/univer) (Apache-2.0).
Single image, single-process: web + Hocuspocus + Fastify in one container.

Source: [github.com/schnsrw/sheets](https://github.com/schnsrw/sheets) · Demo (single-user): [sheet.schnsrw.live](https://sheet.schnsrw.live/)

## Quick start

```sh
# In-memory rooms — fine for a quick try (they vanish on restart):
docker run --rm -p 3000:3000 schnsrw/casual-sheets:latest
# open http://localhost:3000
```

```sh
# With persistence + Redis (recommended):
docker compose up -d
# open http://localhost:3000
```

`docker-compose.yml` (paste this if you don't want to clone the repo):

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

## Co-edit a sheet

1. Open `http://localhost:3000`. **File → Share for co-editing…** to set a password + role and get two share URLs.
2. Paste either URL into another browser. Owner sees joiners in the title-bar avatar stack within ~1 s.
3. Move your selection — peers see your cursor + name follow it. Type — peers see characters appear in your cell live.

Anonymous, no login. Rooms are addressable at `/r/<roomId>`.

## Configuration

| Env var | Default | Meaning |
| --- | --- | --- |
| `PORT` | `3000` | HTTP + WS port |
| `HOST` | `0.0.0.0` | Bind address |
| `REDIS_URL` | _unset_ | If set, Y.Docs persist with a 7-day TTL |
| `ROOM_TTL_MIN` | `15` | Minutes a room stays in memory after the last client leaves |
| `MAX_UPLOAD_MB` | `100` | Cap on multipart + raw-binary uploads (xlsx seed + gzipped snapshot) |

Bundle-time knobs (`VITE_MAX_OPEN_MB`, `VITE_SOFT_WARN_MB`, `VITE_COLLAB_ENABLED`) are baked into the image at build time. Override with `--build-arg` if you build your own.

## Tags

- `latest` — same as the highest semver tag below.
- `0.0.4`, `0.0.3`, `0.0.2` — pinned releases.

Multi-arch — `linux/amd64` + `linux/arm64`.

## What's inside

- **Web** — React + Vite + TypeScript, served statically from the same Fastify server.
- **Collab** — Yjs + Hocuspocus over WebSocket at `/yjs`. Per-mutation op log with `fromCollab` echo-loop prevention.
- **xlsx I/O** — ExcelJS in dedicated Web Workers; main thread never blocks on multi-MB opens.
- **Plugins** — heavy ones (CF, DV, hyperlink, drawing, sort, filter, find-replace) load lazily and eager-load on snapshot inspection.

## License

Apache-2.0. Vendored Univer source keeps its upstream Apache-2.0 license; we don't modify it.
