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

| Var                         | Default                   | Description                                                                                                                                                                                                                                                            |
| --------------------------- | ------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `PORT`                      | `3000`                    | HTTP + WebSocket listen port. The single port the image exposes.                                                                                                                                                                                                       |
| `HOST`                      | `0.0.0.0`                 | Bind address. `0.0.0.0` exposes to the network; `127.0.0.1` keeps it localhost-only.                                                                                                                                                                                   |
| `REDIS_URL`                 | _unset_                   | Redis connection string for Y.Doc persistence (rooms survive server restarts with a 7-day TTL). Unset → in-memory only; rooms vanish on restart. Compose pattern: `redis://redis:6379`.                                                                                |
| `ROOM_TTL_MIN`              | `15`                      | Minutes a room stays in memory after the last client disconnects. Lower → quicker eviction (less idle memory); higher → friendlier reconnect after a short drop.                                                                                                       |
| `MAX_UPLOAD_MB`             | `100`                     | Hard cap on multipart + raw-binary uploads. Bounds the share-room seed (`.xlsx`) and snapshot (gzipped JSON) upload paths. Raise for users with workbooks above this — and bump `VITE_MAX_OPEN_MB` to match so the browser doesn't post something the server will 413. |
| `MAX_ROOMS`                 | `256`                     | Hard cap on concurrent rooms per process. When at cap, `create()` LRU-evicts the oldest evictable room (no password / no seed / no snapshot). If every slot is non-evictable, returns 503 + `retry-after: 60`. Bounds memory under sustained room-creation abuse.      |
| `RATE_LIMIT_ENABLED`        | `true`                    | Master switch for the @fastify/rate-limit plugin. Set to `false` only for load testing or local dev where the bucket would mask real failures. Do not run production with this disabled.                                                                               |
| `RATE_LIMIT_PER_MIN`        | `60`                      | General write-endpoint limit per source IP per minute. Applied to `POST /api/rooms` (cheap, but the easiest abuse target).                                                                                                                                             |
| `UPLOAD_RATE_LIMIT_PER_MIN` | `12`                      | Upload-endpoint limit per source IP per minute. Applied to `POST /api/rooms/:id/seed` and `POST /api/rooms/:id/snapshot` — they take bytes into memory before persisting, so the bucket is tighter.                                                                    |
| `NODE_ENV`                  | `production` _(in image)_ | Standard Node mode. Set by the Dockerfile; rarely overridden.                                                                                                                                                                                                          |

---

## Storage (runtime · landing in v0.1.0)

Reserved keys for the WOPI host-integration MVP. Currently no-op
on `main`; documented here so operators can plan ahead.

| Var                          | Accepted                                           | Description                                                                                                                                        |
| ---------------------------- | -------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------- |
| `CASUAL_STORAGE`             | `memory` _(default)_ · `local` · `s3` · `postgres` | Selects the WOPI backend. `memory` keeps today's no-DB shape. The other three persist workbooks across restarts.                                   |
| `CASUAL_LOCAL_PATH`          | `/data`                                            | Filesystem root when `CASUAL_STORAGE=local`. Mount with `-v ./workbooks:/data`.                                                                    |
| `CASUAL_S3_ENDPOINT`         | _unset_                                            | S3-API endpoint when `CASUAL_STORAGE=s3`. Examples: `https://s3.amazonaws.com`, `http://minio:9000`, `https://<account>.r2.cloudflarestorage.com`. |
| `CASUAL_S3_REGION`           | `us-east-1`                                        | S3 region. Required by AWS S3; safe to keep at default for MinIO / R2 / B2.                                                                        |
| `CASUAL_S3_BUCKET`           | _unset_                                            | Bucket name.                                                                                                                                       |
| `CASUAL_S3_ACCESS_KEY`       | _unset_                                            | S3 access key. Treat as secret.                                                                                                                    |
| `CASUAL_S3_SECRET_KEY`       | _unset_                                            | S3 secret key. Treat as secret.                                                                                                                    |
| `CASUAL_S3_FORCE_PATH_STYLE` | `false`                                            | Set `true` for MinIO and some self-hosted S3 implementations that require path-style addressing.                                                   |
| `CASUAL_PG_URL`              | _unset_                                            | Postgres connection string when `CASUAL_STORAGE=postgres`. Format: `postgres://user:pass@host:port/db`.                                            |

---

## Personal mode (Phase C of #49 — landed 2026-06-06)

Standalone-docker personal account surface. Off by default — operators
opt in by setting `CASUAL_PERSONAL_MODE`. See
[`docs/self-hosting/personal-mode.md`](./self-hosting/personal-mode.md) for
the full setup recipe.

| Var                     | Accepted                                | Description                                                                                                                                                                                                                                                                                                     |
| ----------------------- | --------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `CASUAL_PERSONAL_MODE`  | `none` _(default)_ · `single` · `multi` | `none` keeps today's WOPI-only / anonymous shape. `single` opens signup for the first account then closes it (one personal user per docker). `multi` keeps signup open; first user is admin, every additional user is a plain account. Per-user file isolation is total — admins cannot see other users' files. |
| `CASUAL_USERS_DB_PATH`  | `/data/users.db`                        | SQLite file holding users, sessions, files registry, and profile blobs. Survives restarts via the volume.                                                                                                                                                                                                       |
| `CASUAL_BOOTSTRAP_USER` | _unset_                                 | `<username>:<password>` — when the users table is empty, creates that admin on first boot. Documented escape hatch for upgrading a pre-Phase-C install without clicking through signup. Silently no-ops once any user exists; re-setting the env var does **not** rotate a live password.                       |

CLI password reset (no SMTP plumbing in v1 — by design, see the
`docs/STORAGE_MODES.md` decisions log):

```sh
# Interactive prompt (echo disabled), from inside the container:
docker exec -it <container> casual-sheets-reset-password <username>

# Non-interactive (CI / scripted):
docker exec <container> casual-sheets-reset-password <username> --password='<new>'
```

---

## Networking (runtime · landing in v0.1.0)

Reserved keys for the admin-panel networking surface. Currently no-op
on `main`.

| Var                    | Default                    | Description                                                                                                                                                      |
| ---------------------- | -------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `CASUAL_PUBLIC_ORIGIN` | _detected_                 | The public URL the server should report in redirects, WOPI `BaseFileName`, share-link generation, OG canonical URLs. Example: `https://sheets.acme.example.com`. |
| `CASUAL_CORS_ORIGINS`  | _empty (same-origin only)_ | Comma-separated origins that may call the API. Empty → same-origin only. Example: `https://app.acme.example.com,https://staging.acme.example.com`.               |
| `CASUAL_TRUST_PROXY`   | `loopback`                 | Which proxy hops we accept `X-Forwarded-*` from. `false` to disable; `true` to trust the immediate upstream; a list of IPs / subnets for explicit allowlisting.  |
| `CASUAL_HSTS_MAX_AGE`  | _unset_                    | Emit `Strict-Transport-Security: max-age=<value>` when set. Only set if HTTPS terminates upstream — sending HSTS over HTTP locks users out.                      |

---

## Admin (runtime · v0.1.0)

The admin panel at `/admin` is gated by env-driven credentials. Set
`CASUAL_ADMIN_USERNAME` + `CASUAL_ADMIN_PASSWORD` to enable; the
panel POSTs them to `/api/admin/login` which constant-time-compares
against env and mints a short-lived admin-role JWT for the session.
`CASUAL_JWT_SECRET` must also be set (≥ 16 chars) so the session
token can be signed.

| Var                        | Default                   | Description                                                                                                                                                                                                                      |
| -------------------------- | ------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `CASUAL_ADMIN_USERNAME`    | _unset (panel disabled)_  | Operator login username for the admin panel. Compared constant-time against `/api/admin/login` body.                                                                                                                             |
| `CASUAL_ADMIN_PASSWORD`    | _unset (panel disabled)_  | Operator login password. Treat as secret.                                                                                                                                                                                        |
| `CASUAL_ADMIN_SESSION_TTL` | `3600`                    | Admin session JWT lifetime in seconds.                                                                                                                                                                                           |
| `CASUAL_ADMIN_CONFIG_PATH` | `/data/casual-admin.json` | Filesystem path where the admin panel persists its JSON config (branding, storage, networking, room limits, auth-hook config, base path, webhook subscriptions). Atomically written with mode 0600 — treat the file as a secret. |

### Admin REST endpoints

| Path                | Method | Auth      | Description                                                                       |
| ------------------- | ------ | --------- | --------------------------------------------------------------------------------- |
| `/api/admin/status` | GET    | public    | `{ configured: bool }` — bootstrap probe for the panel.                           |
| `/api/admin/login`  | POST   | env creds | Returns `{ token, ttl_seconds, username }`. Token is an admin-role JWT.           |
| `/api/admin/config` | GET    | admin JWT | Current config; secret fields redacted to `***`.                                  |
| `/api/admin/config` | PUT    | admin JWT | Patch-merge the on-disk config. Secret fields sent as `***` keep the prior value. |

### Admin config — runtime overrides

The on-disk config persisted at `CASUAL_ADMIN_CONFIG_PATH` overrides
env at runtime. Env provides the bootstrap floor; the admin panel
writes win once set. Fields:

- **branding** — `appName`, `accentColor`, `logoUrl`
- **basePath** — reverse-proxy mount path (e.g. `/sheets`); empty
  means served at root. Affects Fastify route prefix + the SPA's
  asset base.
- **storage** — `backend` (`memory` | `local` | `s3` | `postgres`)
  - per-backend creds. Mirrors the `CASUAL_STORAGE` env above.
- **networking** — `publicOrigin`, `corsOrigins`, `trustProxy`,
  `hstsMaxAge`. Mirrors the networking env vars.
- **limits** — `maxRooms`, `maxFileSizeMb`, `roomTtlMin`,
  `maxUsersPerRoom`.
- **auth** — OIDC / SAML / JWT provider config (OIDC + SAML stub
  in v0.1; JWT actively wired).
- **webhooks** — array of `{ name, url, events, secret, enabled }`
  subscriptions. See _Webhooks_ below.

---

## Webhooks (runtime · v0.1.0)

Server-side events fire HTTP POSTs to operator-configured URLs.
HMAC-SHA256 signs the JSON body when a subscription has a `secret`
— receivers verify via the `X-Casual-Signature: sha256=<hex>`
header.

Subscriptions live in `admin-config.webhooks` (writable via the
admin panel). Each subscription:

```json
{
  "name": "audit-log",
  "url": "https://example.com/hooks/casual",
  "events": ["file.saved", "admin.login"],
  "secret": "shh",
  "enabled": true
}
```

Empty `events` array = subscribed to every event.

### Events

| Event                | Fired when                                                             |
| -------------------- | ---------------------------------------------------------------------- |
| `room.created`       | `POST /api/rooms` creates a new room                                   |
| `room.dropped`       | Last client leaves + GC ticks (after `roomTtlMin`)                     |
| `file.uploaded`      | `POST /api/rooms/:id/seed` succeeds                                    |
| `file.saved`         | `POST /wopi/files/:id/contents` succeeds (download → edit → save flow) |
| `file.deleted`       | `DELETE /wopi/files/:id` (admin only)                                  |
| `user.joined`        | New client joins a room                                                |
| `user.left`          | Client disconnects from a room                                         |
| `admin.login`        | Successful `/api/admin/login`                                          |
| `admin.login_failed` | Failed `/api/admin/login`                                              |

### Payload shape

```json
{
  "event": "file.saved",
  "timestamp": "2026-06-01T14:23:09.123Z",
  "payload": {
    "fileId": "wb-q3-budget",
    "size": 12345,
    "version": "1748872989123-abc12345",
    "user": "alice@acme.example"
  }
}
```

Headers on every dispatch:

| Header               | Description                                                                                                     |
| -------------------- | --------------------------------------------------------------------------------------------------------------- |
| `Content-Type`       | `application/json`                                                                                              |
| `User-Agent`         | `CasualSheets-Webhook/0.1`                                                                                      |
| `X-Casual-Event`     | The event name                                                                                                  |
| `X-Casual-Attempt`   | `1` or `2` — see retry policy below                                                                             |
| `X-Casual-Signature` | `sha256=<hex>` when subscription has a secret. Compute the same way to verify: `hmac-sha256(secret, raw_body)`. |

### Retry policy (v0.1)

- Single retry on non-2xx response or network error.
- Retry fires after 5 s.
- After 2 attempts, the dispatch is logged + dropped. v0.2 adds a
  proper queue with exponential back-off + a dead-letter store.

---

## Auth (runtime · JWT access tokens · v0.1.0)

When `CASUAL_JWT_SECRET` is set, every `/wopi/files/*` request must
carry a valid JWT (in `Authorization: Bearer …` or via the
`?access_token=…` query string — WOPI's standard placement). When
unset, WOPI routes fall through to v0.0.x anonymous-by-URL behaviour
— operators opt in to auth by setting the secret.

| Var                      | Default                 | Description                                                                                                                                                  |
| ------------------------ | ----------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `CASUAL_JWT_SECRET`      | _unset (auth disabled)_ | HMAC-SHA256 shared secret used to sign + verify access tokens. Minimum 16 chars; recommend ≥ 32 random bytes (e.g. `openssl rand -hex 32`). Treat as secret. |
| `CASUAL_JWT_DEFAULT_TTL` | `3600`                  | Default token lifetime in seconds when `ttl_seconds` is omitted from the `POST /api/tokens` body.                                                            |

### Token claims (signed payload)

| Claim               | Type                                             | Description                                                                                                               |
| ------------------- | ------------------------------------------------ | ------------------------------------------------------------------------------------------------------------------------- |
| `sub`               | string                                           | Username, email, or any stable user identifier. Surfaces as `UserId` in CheckFileInfo.                                    |
| `file_id`           | string                                           | The single file this token authorises. WOPI routes reject when the URL `:id` ≠ this claim.                                |
| `role`              | `'admin' \| 'editor' \| 'commenter' \| 'viewer'` | Coarse role. Default permission map applied unless `permissions` overrides.                                               |
| `permissions`       | object _(optional)_                              | Per-flag override: `{ read, write, comment, download, share, admin }`.                                                    |
| `features`          | object _(optional)_                              | Feature toggles consumed by the client UI: `{ charts, pivots, conditionalFormatting, sharing, exportFiles, collab, ai }`. |
| `password_required` | boolean _(optional)_                             | When true, the legacy `x-room-password` header gate also applies on top of the JWT.                                       |
| `display_name`      | string _(optional)_                              | Human label for presence + cursor markers. Falls back to `sub`.                                                           |
| `aud`               | string _(optional)_                              | Audience — typically the deployment's public origin.                                                                      |
| `iss`               | string _(optional)_                              | Issuer — useful when downstream SSO mints tokens.                                                                         |
| `exp`, `iat`        | number                                           | Standard JWT expiry + issued-at, set by the signer.                                                                       |

### Endpoints exposed when JWT is configured

| Path          | Method | Auth                | Description                                                                                                                                                                                  |
| ------------- | ------ | ------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `/api/me`     | GET    | optional            | Decodes the token + returns resolved role / permissions / features / `passwordRequired` / `exp`. Returns `{ anonymous: true }` without a token.                                              |
| `/api/tokens` | POST   | admin role required | Mint a new token. Body: `{ sub, file_id, role, permissions?, features?, password_required?, display_name?, ttl_seconds?, aud?, iss? }`. Returns the signed JWT + the resolved claim summary. |
| `/api/files`  | GET    | admin role required | List every file id the host backend knows about.                                                                                                                                             |

### Bootstrapping the first admin token

Tokens are minted by tokens — chicken-and-egg. Sign the first admin
token manually using the secret:

```sh
docker compose exec app node -e '
  import("jsonwebtoken").then(({ default: jwt }) => {
    const tok = jwt.sign(
      { sub: "owner", file_id: "*", role: "admin" },
      process.env.CASUAL_JWT_SECRET,
      { algorithm: "HS256", expiresIn: "8h" },
    );
    console.log(tok);
  });
'
```

After that, the admin panel (v0.1.0) issues subsequent tokens through
its UI; the bootstrap step happens once per deployment.

---

## Web build (Vite — bake-time only)

These are read at `pnpm --filter @sheet/web build` and bundled into
the JS. Override via `--build-arg`. Setting them at runtime does
nothing.

| Var                   | Default              | Description                                                                                                                                                   |
| --------------------- | -------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `VITE_COLLAB_ENABLED` | `1`                  | Ship co-editing in the bundle. Off in the GitHub Pages demo build; on in the Docker image.                                                                    |
| `VITE_COLLAB_WS_URL`  | _same-origin `/yjs`_ | WebSocket URL the collab driver dials. Override when running Vite dev (`:5273`) against a standalone server (`:3000`).                                        |
| `VITE_MAX_OPEN_MB`    | `100`                | Hard reject for File → Open / drag-drop. Larger files freeze and eventually OOM-crash the tab during the ExcelJS parse. The supported sweet spot is 25–50 MB. |
| `VITE_SOFT_WARN_MB`   | `25`                 | Threshold above which the loading overlay shows the up-front "this is a large workbook, may take 10+ s" hint. Should be ≤ `VITE_MAX_OPEN_MB`.                 |

---

## OCI image-label build args

Passed by `.github/workflows/docker-publish.yml` at tag-time. Surface
as `org.opencontainers.image.*` labels on the published image so
operators can `docker inspect` provenance.

| Build arg           | Sets label       | Notes                                 |
| ------------------- | ---------------- | ------------------------------------- |
| `CASUAL_VERSION`    | `image.version`  | The git tag, e.g. `v0.1.0`.           |
| `CASUAL_GIT_SHA`    | `image.revision` | Full commit SHA at the tag.           |
| `CASUAL_BUILD_DATE` | `image.created`  | RFC 3339 UTC timestamp at build time. |

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
