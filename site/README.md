# schnsrw.live — apex home page

Static landing page that serves the apex of `schnsrw.live`. Pure HTML +
CSS, no build step. The sheet app stays at `sheet.schnsrw.live` (built
from `apps/web/` in this repo); this site is a separate deployment.

## Why a separate site

A single GitHub Pages site = one repo. The sheet app already owns this
repo's Pages deploy (`sheet.schnsrw.live`). The apex needs its own
Pages site, which means either a second repo or a different host.

## Deploying to GitHub Pages (recommended)

1. Create a new repo, e.g. `schnsrw/schnsrw.github.io` (the name is
   significant if you use GitHub's user-site pattern — it gets auto-
   served at `https://schnsrw.github.io/`).
2. Copy the contents of this `site/` directory into the new repo's
   root (including `CNAME` and `favicon.svg`).
3. Push to `main`.
4. In the new repo: Settings → Pages → Source: "Deploy from a branch"
   → branch `main`, folder `/ (root)`.
5. Still in Settings → Pages: set the custom domain to `schnsrw.live`
   and tick "Enforce HTTPS" (allow a few minutes for the cert).
6. DNS at your registrar — point the apex to GitHub:

   ```
   A    @    185.199.108.153
   A    @    185.199.109.153
   A    @    185.199.110.153
   A    @    185.199.111.153
   ```

   And make sure `sheet.schnsrw.live` already has its CNAME pointing
   to `schnsrw.github.io.` for the sheet app.

## Deploying elsewhere

Drop `site/` onto any static host — Vercel, Netlify, Cloudflare Pages,
Render. Point them at this directory; no build command needed.

## Local preview

```sh
cd site
python3 -m http.server 4000
# open http://localhost:4000
```
