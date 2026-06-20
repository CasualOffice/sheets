# syntax=docker/dockerfile:1.7
#
# Casual Sheets — single image, web + server.
#
#   docker build -t casualoffice/sheets:latest .
#   docker run -p 3000:3000 casualoffice/sheets:latest
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

# The fork's `link:` overrides in package.json resolve to
# vendor/univer-revamp/packages/<name>. The host CI runs
# `./scripts/setup-fork.sh` BEFORE `docker compose up --build`,
# which builds lib/ + swaps each fork package.json's main/exports
# from the dev shape to the consumable shape. We bring in the
# fork artifacts at the same paths so the override links resolve.
# `.dockerignore` strips src/, node_modules/, and other dev-only
# bits from the build context to keep the image small.
COPY vendor/univer-revamp vendor/univer-revamp

COPY apps/web/package.json apps/web/
COPY apps/server/package.json apps/server/
# packages/sdk is in the pnpm workspace too — apps/web now imports
# @casualoffice/sheets/xlsx (and friends) via the workspace symlink.
# Without these copies pnpm install fails the lockfile check; without
# the source the SDK build below has nothing to compile.
COPY packages/sdk/package.json packages/sdk/

# Full install (dev + prod). The build stage needs Vite & TypeScript; we
# trim back to prod-only deps in the runtime stage below.
RUN --mount=type=cache,id=pnpm,target=/root/.local/share/pnpm/store \
    pnpm install --frozen-lockfile

# ─────────────── build-web ───────────────
FROM deps AS build-web
COPY packages/sdk packages/sdk
COPY apps/web apps/web

# Build the SDK first so apps/web's tsc + vite resolve
# @casualoffice/sheets/* through packages/sdk/dist/. Skipping this
# is what caused the e2e-prod failure on 4dc4dc4 — host CI builds the
# SDK separately but the Docker image's pnpm sandbox doesn't know to.
RUN pnpm --filter @casualoffice/sheets build

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
# Bump the V8 heap: the web bundle is large (Univer + the SDK editor core,
# which apps/web now mounts via `<CasualSheets>`), and the default heap OOMs
# the Vite build (SIGABRT) in the constrained Docker VM. Builds fine locally;
# this just gives CI/Docker the same headroom.
ENV NODE_OPTIONS=--max-old-space-size=4096
RUN pnpm --filter @sheet/web build

# ─────────────── runtime ───────────────
FROM node:${NODE_VERSION} AS runtime
RUN corepack enable && corepack prepare pnpm@10.33.4 --activate

# wget for HEALTHCHECK (alpine has it via busybox). curl would also work.
# su-exec drops root → node in the entrypoint after fixing data-dir ownership.
RUN apk add --no-cache wget su-exec

ENV NODE_ENV=production \
    HOST=0.0.0.0 \
    PORT=3000

WORKDIR /app

# Copy lockfile + manifests + workspace config — required for `pnpm install
# --prod` to resolve the workspace graph cleanly.
COPY pnpm-lock.yaml pnpm-workspace.yaml package.json ./

# Same fork copy as in the deps stage — runtime needs the package.jsons
# + lib/ outputs at vendor/univer-revamp/packages/* so the override
# `link:` paths resolve at server start.
COPY vendor/univer-revamp vendor/univer-revamp

COPY apps/web/package.json apps/web/
COPY apps/server/package.json apps/server/
# Same workspace-manifest copy as in deps — pnpm install --prod resolves
# the workspace graph and needs every package.json present even if the
# runtime image never executes the SDK directly (apps/web's dist/ is
# already baked in by the build-web stage).
COPY packages/sdk/package.json packages/sdk/

# Prod-only install. tsx lives in dependencies (not devDeps) so the server
# can run TypeScript directly without a JS compile step.
RUN --mount=type=cache,id=pnpm,target=/root/.local/share/pnpm/store \
    pnpm install --frozen-lockfile --prod

# Bring in the server source + the built web bundle. Server source is small
# and tsx executes it directly at startup.
COPY apps/server/src apps/server/src
COPY apps/server/tsconfig.json apps/server/
COPY --from=build-web /repo/apps/web/dist apps/web/dist

# ─────────────── OCI image labels ───────────────
#
# Standard `org.opencontainers.image.*` keys baked into every published
# image. The CI workflow passes these as build args at tag-time:
#
#     --build-arg CASUAL_VERSION=v0.1.0   (the git tag)
#     --build-arg CASUAL_GIT_SHA=abc1234  (full commit SHA)
#     --build-arg CASUAL_BUILD_DATE=...   (RFC 3339 UTC timestamp)
#
# Inspect with:
#     docker inspect casualoffice/sheets:latest | jq '.[0].Config.Labels'
#
# Downstream operators rely on these labels to pin a specific build;
# they're the documentation that travels with the artifact.
ARG CASUAL_VERSION=dev
ARG CASUAL_GIT_SHA=unknown
ARG CASUAL_BUILD_DATE=unknown

LABEL org.opencontainers.image.title="Casual Sheets" \
      org.opencontainers.image.description="Excel-flavored web spreadsheet with real-time co-editing. Single image: web app, Hocuspocus WebSocket gateway, and Fastify HTTP server on one port. Built on Univer OSS." \
      org.opencontainers.image.url="https://sheet.casualoffice.org/" \
      org.opencontainers.image.source="https://github.com/CasualOffice/sheets" \
      org.opencontainers.image.documentation="https://casualoffice.org/docs/sheets/" \
      org.opencontainers.image.vendor="Sachin Sarwa" \
      org.opencontainers.image.authors="Sachin Sarwa <schnsrw@gmail.com>" \
      org.opencontainers.image.licenses="Apache-2.0" \
      org.opencontainers.image.version="${CASUAL_VERSION}" \
      org.opencontainers.image.revision="${CASUAL_GIT_SHA}" \
      org.opencontainers.image.created="${CASUAL_BUILD_DATE}"

EXPOSE 3000

HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 \
    CMD wget -q --spider http://127.0.0.1:3000/health || exit 1

# Provision the personal-mode storage root with `node` ownership at build time.
# /data is the default CASUAL_LOCAL_PATH. This covers named volumes (Docker
# copies the image's ownership onto an empty volume), but NOT bind mounts — a
# host dir bind-mounted at /data keeps its host ownership (root, if Docker
# created it), masking this. The entrypoint below fixes that case at runtime.
RUN mkdir -p /data && chown -R node:node /data

# Runtime ownership fix + privilege drop. The container starts as ROOT so the
# entrypoint can chown the (possibly root-owned, bind-mounted) data dir, then
# drops to `node` via su-exec to run the server. Without starting as root we
# couldn't repair a bind-mounted /data and personal mode would crash with
# SQLITE_CANTOPEN (GitHub #57).
COPY scripts/docker-entrypoint.sh /usr/local/bin/docker-entrypoint.sh
RUN chmod +x /usr/local/bin/docker-entrypoint.sh

WORKDIR /app/apps/server
ENTRYPOINT ["/usr/local/bin/docker-entrypoint.sh"]
CMD ["node", "--import", "tsx", "src/index.ts"]
