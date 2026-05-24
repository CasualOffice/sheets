# Environment variables

The single source of truth for every runtime + build-time knob Casual
Sheets reads. The admin panel (lands in v0.1.0) reads this doc to
render its config UI; `docs/DOCKERHUB.md` and the site self-hosting
section both link here.

Two flavours:

- **Runtime** — read by the Node server at startup or per-request.
  Settable via `docker run -e`, the `environment:` block in
  `docker-compose.yml`, or `--env-file`.
- **Build-time** — read by Vite during `pnpm --filter @sheet/web
  build` and baked into the frontend bundle. Setting these at runtime
  does nothing; pass them as `--build-arg` on `docker build` or via
  the `args:` block of `docker-compose.yml` to bake your own image.

---

## Server (runtime)

| Var | Default | Description |
|---|---|---|
| `PORT` | `3000` | HTTP + WebSocket listen port. The single port the image exposes. |
| `HOST` | `0.0.0.0` | Bind address. `0.0.0.0` exposes to the network; `127.0.0.1` keeps it localhost-only. |
| `REDIS_URL` | _unset_ | Redis connection string for Y.Doc persistence (rooms survive server restarts with a 7-day TTL). Unset → in-memory only; rooms vanish on restart. Compose pattern: `redis://redis:6379`. |
| `ROOM_TTL_MIN` | `15` | Minutes a room stays in memory after the last client disconnects. Lower → quicker eviction (less idle memory); higher → friendlier reconnect after a short drop. |
| `MAX_UPLOAD_MB` | `100` | Hard cap on multipart + raw-binary uploads. Bounds the share-room seed (`.xlsx`) and snapshot (gzipped JSON) upload paths. Raise for users with workbooks above this — and bump `VITE_MAX_OPEN_MB` to match so the browser doesn't post something the server will 413. |
| `NODE_ENV` | `production` _(in image)_ | Standard Node mode. Set by the Dockerfile; rarely overridden. |

---

## Storage (runtime · landing in v0.1.0)

Reserved keys for the WOPI host-integration MVP. Currently no-op
on `main`; documented here so operators can plan ahead.

| Var | Accepted | Description |
|---|---|---|
| `CASUAL_STORAGE` | `memory` _(default)_ · `local` · `s3` · `postgres` | Selects the WOPI backend. `memory` keeps today's no-DB shape. The other three persist workbooks across restarts. |
| `CASUAL_LOCAL_PATH` | `/data` | Filesystem root when `CASUAL_STORAGE=local`. Mount with `-v ./workbooks:/data`. |
| `CASUAL_S3_ENDPOINT` | _unset_ | S3-API endpoint when `CASUAL_STORAGE=s3`. Examples: `https://s3.amazonaws.com`, `http://minio:9000`, `https://<account>.r2.cloudflarestorage.com`. |
| `CASUAL_S3_REGION` | `us-east-1` | S3 region. Required by AWS S3; safe to keep at default for MinIO / R2 / B2. |
| `CASUAL_S3_BUCKET` | _unset_ | Bucket name. |
| `CASUAL_S3_ACCESS_KEY` | _unset_ | S3 access key. Treat as secret. |
| `CASUAL_S3_SECRET_KEY` | _unset_ | S3 secret key. Treat as secret. |
| `CASUAL_S3_FORCE_PATH_STYLE` | `false` | Set `true` for MinIO and some self-hosted S3 implementations that require path-style addressing. |
| `CASUAL_PG_URL` | _unset_ | Postgres connection string when `CASUAL_STORAGE=postgres`. Format: `postgres://user:pass@host:port/db`. |

---

## Networking (runtime · landing in v0.1.0)

Reserved keys for the admin-panel networking surface. Currently no-op
on `main`.

| Var | Default | Description |
|---|---|---|
| `CASUAL_PUBLIC_ORIGIN` | _detected_ | The public URL the server should report in redirects, WOPI `BaseFileName`, share-link generation, OG canonical URLs. Example: `https://sheets.acme.example.com`. |
| `CASUAL_CORS_ORIGINS` | _empty (same-origin only)_ | Comma-separated origins that may call the API. Empty → same-origin only. Example: `https://app.acme.example.com,https://staging.acme.example.com`. |
| `CASUAL_TRUST_PROXY` | `loopback` | Which proxy hops we accept `X-Forwarded-*` from. `false` to disable; `true` to trust the immediate upstream; a list of IPs / subnets for explicit allowlisting. |
| `CASUAL_HSTS_MAX_AGE` | _unset_ | Emit `Strict-Transport-Security: max-age=<value>` when set. Only set if HTTPS terminates upstream — sending HSTS over HTTP locks users out. |

---

## Admin (runtime · landing in v0.1.0)

| Var | Default | Description |
|---|---|---|
| `CASUAL_ADMIN_PASSWORD` | _unset (panel disabled)_ | Required to unlock the admin panel at `/admin`. v0.1 ships single-admin auth; v0.2 will add proper admin accounts. |
| `CASUAL_ADMIN_CONFIG_PATH` | `/data/casual-admin.json` | Filesystem path where the admin panel persists its JSON config (branding, storage, networking, room limits, auth-hook config). |

---

## Web build (Vite — bake-time only)

These are read at `pnpm --filter @sheet/web build` and bundled into
the JS. Override via `--build-arg`. Setting them at runtime does
nothing.

| Var | Default | Description |
|---|---|---|
| `VITE_COLLAB_ENABLED` | `1` | Ship co-editing in the bundle. Off in the GitHub Pages demo build; on in the Docker image. |
| `VITE_COLLAB_WS_URL` | _same-origin `/yjs`_ | WebSocket URL the collab driver dials. Override when running Vite dev (`:5273`) against a standalone server (`:3000`). |
| `VITE_MAX_OPEN_MB` | `100` | Hard reject for File → Open / drag-drop. Larger files freeze and eventually OOM-crash the tab during the ExcelJS parse. The supported sweet spot is 25–50 MB. |
| `VITE_SOFT_WARN_MB` | `25` | Threshold above which the loading overlay shows the up-front "this is a large workbook, may take 10+ s" hint. Should be ≤ `VITE_MAX_OPEN_MB`. |

---

## OCI image-label build args

Passed by `.github/workflows/docker-publish.yml` at tag-time. Surface
as `org.opencontainers.image.*` labels on the published image so
operators can `docker inspect` provenance.

| Build arg | Sets label | Notes |
|---|---|---|
| `CASUAL_VERSION` | `image.version` | The git tag, e.g. `v0.1.0`. |
| `CASUAL_GIT_SHA` | `image.revision` | Full commit SHA at the tag. |
| `CASUAL_BUILD_DATE` | `image.created` | RFC 3339 UTC timestamp at build time. |

Inspect with:

```sh
docker inspect schnsrw/casual-sheets:latest \
  | jq '.[0].Config.Labels | with_entries(select(.key | startswith("org.opencontainers")))'
```

---

## Discovery convention

- **Server** keys use snake-case (`PORT`, `ROOM_TTL_MIN`) for backwards
  compatibility with the v0.0.x release line.
- **Storage / networking / admin** keys (v0.1.0+) all carry the
  `CASUAL_` prefix so they're greppable and don't collide with
  generic env vars on a shared host.

If you're adding a new runtime knob, follow the `CASUAL_*` convention
and update this file in the same commit. The admin panel auto-renders
its config UI from this table.
