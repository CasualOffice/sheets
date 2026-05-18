# syntax=docker/dockerfile:1.7
#
# Casual Sheets — single image, web + server.
#
#   docker build -t schnsrw/casual-sheets:latest .
#   docker run -p 3000:3000 schnsrw/casual-sheets:latest
#   open http://localhost:3000
#
# For persistence (rooms survive restarts), wire Redis via compose:
#   docker compose up
#
# Multi-stage layout:
#   deps         — installs the workspace's full dep graph once
#   build-web    — produces apps/web/dist (served statically by the server)
#   runtime      — node:22-alpine + tsx + only what the server needs at runtime
#
# Pinned Node 22-alpine for a small base; pnpm version pulled from
# package.json's `packageManager` field so it matches the workspace.

ARG NODE_VERSION=22-alpine

# ─────────────── deps ───────────────
FROM node:${NODE_VERSION} AS deps
RUN corepack enable && corepack prepare pnpm@10.33.4 --activate
WORKDIR /repo

# Copy lockfile + workspace manifest first so Docker can cache the dep
# install across source-only changes.
COPY pnpm-lock.yaml pnpm-workspace.yaml package.json ./
COPY patches patches
COPY apps/web/package.json apps/web/
COPY apps/server/package.json apps/server/

# Full install (dev + prod). The build stage needs Vite & TypeScript; we
# trim back to prod-only deps in the runtime stage below.
RUN --mount=type=cache,id=pnpm,target=/root/.local/share/pnpm/store \
    pnpm install --frozen-lockfile

# ─────────────── build-web ───────────────
FROM deps AS build-web
COPY apps/web apps/web

# Build-time knobs the Vite bundle bakes in. Override via
# `--build-arg VITE_MAX_OPEN_MB=200` on `docker build` or via the
# `args:` block of `docker-compose.yml`. Defaults match `.env.example`.
ARG VITE_COLLAB_ENABLED=1
ARG VITE_MAX_OPEN_MB=100
ARG VITE_SOFT_WARN_MB=25
ENV VITE_COLLAB_ENABLED=${VITE_COLLAB_ENABLED}
ENV VITE_MAX_OPEN_MB=${VITE_MAX_OPEN_MB}
ENV VITE_SOFT_WARN_MB=${VITE_SOFT_WARN_MB}

# Same base path the production deploy expects — assets resolve from the
# server's root. PAGES_BASE only matters for the GitHub Pages build.
RUN pnpm --filter @sheet/web build

# ─────────────── runtime ───────────────
FROM node:${NODE_VERSION} AS runtime
RUN corepack enable && corepack prepare pnpm@10.33.4 --activate

# wget for HEALTHCHECK (alpine has it via busybox). curl would also work.
RUN apk add --no-cache wget

ENV NODE_ENV=production \
    HOST=0.0.0.0 \
    PORT=3000

WORKDIR /app

# Copy lockfile + manifests + workspace config — required for `pnpm install
# --prod` to resolve the workspace graph cleanly.
COPY pnpm-lock.yaml pnpm-workspace.yaml package.json ./
COPY patches patches
COPY apps/web/package.json apps/web/
COPY apps/server/package.json apps/server/

# Prod-only install. tsx lives in dependencies (not devDeps) so the server
# can run TypeScript directly without a JS compile step.
RUN --mount=type=cache,id=pnpm,target=/root/.local/share/pnpm/store \
    pnpm install --frozen-lockfile --prod

# Bring in the server source + the built web bundle. Server source is small
# and tsx executes it directly at startup.
COPY apps/server/src apps/server/src
COPY apps/server/tsconfig.json apps/server/
COPY --from=build-web /repo/apps/web/dist apps/web/dist

EXPOSE 3000

HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 \
    CMD wget -q --spider http://127.0.0.1:3000/health || exit 1

# Drop privileges. node:alpine ships a `node` user.
USER node

WORKDIR /app/apps/server
CMD ["node", "--import", "tsx", "src/index.ts"]
