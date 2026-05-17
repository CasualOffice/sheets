# schnsrw.live — apex home page

Static landing page bundled into this repo's GitHub Pages deploy and
served at the apex of `schnsrw.live`. Pure HTML + CSS, no build step.

## How it ships

The `.github/workflows/deploy-pages.yml` workflow assembles a single
Pages artifact from two sources:

| URL                       | Source                  |
|---------------------------|-------------------------|
| `schnsrw.live/`           | `site/`                 |
| `schnsrw.live/sheets/`    | `apps/web/dist/` (built with `PAGES_BASE=/sheets/`) |

The artifact root carries a `CNAME` file (`schnsrw.live`) that tells
GitHub Pages which host to serve. Both the sheet app and the landing
page deploy together on every push to `main`.

The GitHub-provided URL `https://schnsrw.github.io/sheets/` keeps
working (it 301s to `schnsrw.live/sheets/`).

## DNS at your registrar

Apex (`schnsrw.live`) — four A records:

```
A    @    185.199.108.153
A    @    185.199.109.153
A    @    185.199.110.153
A    @    185.199.111.153
```

If you also want a `sheet.schnsrw.live` subdomain pointing at the same
app, the cleanest setup is a 301 redirect at your registrar
(Cloudflare Page Rule, Namecheap URL Redirect, etc.) sending
`sheet.schnsrw.live/*` → `https://schnsrw.live/sheets/*`.

## Local preview

```sh
cd site
python3 -m http.server 4000
# open http://localhost:4000
```
