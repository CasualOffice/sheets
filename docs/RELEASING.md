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
| Release tag form | `v0.3.2` | `@casualoffice/sheets@0.8.0` |
| Shows up as | a Docker Hub tag | an npm version **+ a GitHub Release** |
| **Latest** | **0.3.2** | **0.8.0** |

> The **GitHub Releases page tracks the SDK line** (Changesets creates a release
> per npm publish). The **Docker image follows the `vX.Y.Z` git tags**. So the
> Releases page reading `0.8.0` while the Docker image is `0.3.2` is expected —
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

The project is mid-rename from the `schnsrw` scope to `casualoffice`. The Docker
line has finished migrating; the npm line has not:

- **Docker app — latest `0.3.2`, published under `casualoffice/sheets`.**
  (`:0.3.2`, `:0.3`, `:0`, `:latest` all live.) The old `schnsrw/casual-sheets`
  tags (`:0.3.1` / `:latest`) remain live but are frozen. There is no `0.4.x` /
  `0.5.x` Docker image — those numbers belong to the SDK line, not the app
  (see [#57](https://github.com/CasualOffice/sheets/issues/57)).
- **SDK — two generations.** The **old** line is published as
  `@schnsrw/casual-sheets@0.8.0` (pre-restructure: minimal editor + xlsx import +
  host-controlled toolbar). The **new** line is `@casualoffice/sheets` — the
  Excalidraw-model restructure now landing on `main` (full editor, formula
  engine, `CasualSheetsAPI`, `onChange`, lazy plugins, …). It lives in
  `packages/sdk` but is **not yet published** (the scope 404s); the unreleased
  changesets ship it as `@casualoffice/sheets@0.9.0+` on first publish. The new
  API does **not** exist in `@schnsrw/casual-sheets@0.8.0`.

**Until the `@casualoffice/sheets` publish lands:** the restructure SDK is
install-able only from source; `@schnsrw/casual-sheets@0.8.0` is the only
published SDK but lacks the new API. Docs reference the going-forward name
`@casualoffice/sheets`. Self-hosters pull `casualoffice/sheets:latest`.
