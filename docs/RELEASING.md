# Releasing

This repo ships **two independent artifacts on two independent version lines.**
They serve different consumers and move at different cadences, so their version
numbers do **not** track each other. This is the #1 source of "which version is
the project at?" confusion — read this before cutting or citing a release.

| | **App / Docker image** | **SDK / npm package** |
| --- | --- | --- |
| Artifact | `casualoffice/sheets` Docker image | `@casualoffice/sheets` npm package |
| What it is | the self-hostable spreadsheet (server + web) | the embeddable editor for integrators |
| Lives in | `apps/web` + `apps/server` | `packages/sdk` |
| Versioned by | `vX.Y.Z` **git tags** | **Changesets** |
| Workflow | `.github/workflows/docker-publish.yml` | `.github/workflows/release-npm.yml` |
| Release tag form | `v0.3.1` | `@casualoffice/sheets@0.8.0` |
| Shows up as | a Docker Hub tag | an npm version **+ a GitHub Release** |
| **Latest** | **0.3.1** | **0.8.0** |

> The **GitHub Releases page tracks the SDK line** (Changesets creates a release
> per npm publish). The **Docker image follows the `vX.Y.Z` git tags**. So the
> Releases page reading `0.8.0` while the Docker image is `0.3.1` is expected —
> they are different things.

---

## Cutting an App / Docker release

1. Bump the version in the root `package.json` (this is the app/Docker line).
2. Tag and push:
   ```bash
   git tag v0.3.2 && git push origin v0.3.2
   ```
3. The tag push triggers `docker-publish.yml`, which builds and pushes
   `casualoffice/sheets` as `:{version}`, `:{major}.{minor}`, `:{major}`, and
   `:latest`. (Manual re-run: `gh workflow run docker-publish.yml -f tag=v0.3.2`.)

## Cutting an SDK / npm release

1. Add a changeset describing the change:
   ```bash
   pnpm changeset
   ```
2. Merge the Changesets "Version Packages" PR (or run `release-npm.yml`). That
   builds `packages/sdk` and runs `changeset publish` → npm publish under the
   `@casualoffice` scope + a GitHub Release named `@casualoffice/sheets@x.y.z`.

The two lines are deliberately decoupled: an SDK fix does **not** require a new
Docker image, and a Docker app release does **not** require an npm bump.

---

## Current state & in-progress cleanup (2026-06)

The project is mid-rename from the `schnsrw` scope to `casualoffice`; until it
settles, registries are split:

- **Docker app — latest `0.3.1`.** Historically published as
  `schnsrw/casual-sheets` (`:0.3.1` / `:latest` still live and pullable).
  Migrating to `casualoffice/sheets`; the **next release publishes under
  `casualoffice/sheets`**. There is no `0.4.x` / `0.5.x` Docker image — those
  numbers belong to the SDK line, not the app (see [#57](https://github.com/CasualOffice/sheets/issues/57)).
- **SDK — latest `0.8.0`.** Currently on npm under the old scope
  `@schnsrw/casual-sheets`. The first `@casualoffice/sheets` npm publish is
  **pending** (the new scope 404s until then).

**Until the migration completes:** self-hosters should pull
`schnsrw/casual-sheets:latest`; integrators should depend on
`@schnsrw/casual-sheets` until `@casualoffice/sheets` is published.
