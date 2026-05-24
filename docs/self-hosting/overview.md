# Self-hosting overview

Casual Sheets ships as a single Docker image:

```sh
docker run --rm -p 3000:3000 schnsrw/casual-sheets:0.1
# open http://localhost:3000
```

Inside the container: the React web app, the Hocuspocus WebSocket
gateway, and the Fastify HTTP server all run on **one port**. No
reverse proxy is required to start; you can add one when you want
TLS / a custom domain / multi-app multiplexing.

## Three deployment shapes

How big is "production" for you? Pick the shape:

| Shape | What | When |
|---|---|---|
| **1 — Local dev / personal** | `docker run -p 3000:3000` | Try it. Single host. No persistence — rooms vanish on restart. |
| **2 — Single-host, persistent** | Docker Compose with Redis + a host volume + your reverse proxy | A team or department. Sticky sessions, but no horizontal scale. |
| **3 — Multi-host (v0.2 lane)** | Multiple replicas behind a load balancer + S3 / Postgres workbook storage + Redis pub/sub for room state | A real org. Horizontal WebSocket scaling has open questions documented below. |

This page covers shapes 1 + 2. Shape 3 has _shape-changing_ open
items that v0.2 closes (sticky WebSocket session handling for the
y-websocket protocol across replicas; shared awareness backplane).
The pieces work in isolation today — the integration is the bit
that needs hardening.

## What you'll find in this section

- **[Reverse proxy recipes](./reverse-proxy.md)** — nginx, Caddy,
  Traefik. The WebSocket upgrade header is the most-common
  gotcha; recipes include the minimum config + commentary.
- **[TLS + custom domain](./tls.md)** — Let's Encrypt with each
  proxy, DNS pointers, the `CASUAL_PUBLIC_ORIGIN` setting.
- **[CORS](./cors.md)** — when you need it, when you don't, the
  `CASUAL_CORS_ORIGINS` env var, the most common mistake.
- **[Scaling](./scaling.md)** — single-process limits today, Redis
  persistence, horizontal scale-out caveats labelled as v0.2.
- **[Backups](./backups.md)** — per-backend recipes (S3 versioning,
  Postgres `pg_dump`, local rsync). Restore drill.

## Configuration model

Three layers of configuration, in precedence order:

1. **On-disk config** at `CASUAL_ADMIN_CONFIG_PATH` (default
   `/data/casual-admin.json`). Written by the admin panel.
   Reflected at runtime — no restart needed for most changes.
2. **Env vars** (the `CASUAL_*` set). Bootstrap floor. Operators
   ship a deployment with sensible defaults via env; the admin
   panel overrides individual fields.
3. **Compiled defaults**. Last resort. See
   [`apps/server/src/admin/config.ts`](https://github.com/schnsrw/sheets/blob/main/apps/server/src/admin/config.ts).

Full env-var matrix: [`docs/ENV.md`](../ENV.md).

## Single-host, persistent — `docker compose`

```yaml
services:
  app:
    image: schnsrw/casual-sheets:0.1
    restart: unless-stopped
    ports: ['3000:3000']
    environment:
      # ── Persistence ──
      REDIS_URL: redis://redis:6379

      # ── Workbook storage (file persistence) ──
      CASUAL_STORAGE: local
      CASUAL_LOCAL_PATH: /data/workbooks

      # ── Admin panel ──
      CASUAL_ADMIN_USERNAME: admin
      CASUAL_ADMIN_PASSWORD: ${ADMIN_PASSWORD}   # from .env file
      CASUAL_JWT_SECRET: ${JWT_SECRET}           # 32+ random chars
      CASUAL_ADMIN_CONFIG_PATH: /data/casual-admin.json

      # ── Networking ──
      CASUAL_PUBLIC_ORIGIN: https://sheets.acme.example
    volumes:
      - data:/data
    depends_on:
      redis: { condition: service_healthy }

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
  data:
  redis-data:
```

Generate the secrets once:

```sh
echo "ADMIN_PASSWORD=$(openssl rand -hex 16)" >> .env
echo "JWT_SECRET=$(openssl rand -hex 32)" >> .env
```

Then `docker compose up -d` + visit `https://sheets.acme.example/admin`
to finish wiring branding, room limits, webhooks etc. from the panel.

## What happens on every restart

- **Y.Doc room state** — survives via Redis (`REDIS_URL`). 7-day
  idle TTL.
- **Workbook files** — survive via `CASUAL_STORAGE=local` writing to
  the `/data` volume (or via `s3` / `postgres` if you wired one of
  those).
- **Admin config** — survives via `CASUAL_ADMIN_CONFIG_PATH`
  writing to `/data/casual-admin.json`. Atomically updated; mode
  `0600` (treat as secret).
- **In-flight WebSocket sessions** — don't survive. Clients
  reconnect, the room state replays from Redis, no data loss.

## Where to go next

- [`docs/customization/`](../customization/) — admin-panel walkthrough,
  per-field reference, webhook signature verification examples, white-
  labelling.
- [`docs/ENV.md`](../ENV.md) — every env var, every default, every
  related route.
- [`docs/CO-EDITING.md`](../CO-EDITING.md) — protocol details for
  the y-websocket layer (read this before scaling beyond one
  replica).
