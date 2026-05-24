# TLS + custom domain

TLS termination happens at the reverse proxy. Casual Sheets itself
listens on plain HTTP inside the container; that's the standard
shape for self-hosted services and keeps cert renewal cleanly
isolated from the application lifecycle.

## DNS

Point an `A` (or `AAAA`) record at your server's public IP:

```
sheets.acme.example.   3600    IN  A   203.0.113.42
```

Or — if you're under a CDN like Cloudflare with a proxied record —
the record can be `CNAME` to the Cloudflare hostname and Cloudflare
terminates the TLS.

## Let's Encrypt

Each proxy has its own automation:

- **Caddy** — automatic. First request to the hostname triggers an
  HTTP-01 challenge; cert provisioned + renewed in the background.
- **Traefik** — `certificatesresolvers.letsencrypt.acme.*` block in
  the static config (see `reverse-proxy.md`).
- **nginx** — `certbot --nginx -d sheets.acme.example`. Cron / systemd
  timer handles renewal.

For all three: TCP/80 must be reachable from the public internet
during the challenge (or use DNS-01 if you have API access to your
DNS provider — Caddy + Traefik both support most DNS providers via
plugins).

## `CASUAL_PUBLIC_ORIGIN`

Once HTTPS is up, set the public origin so the server emits
correct redirect URLs + WOPI `BaseFileName` values:

```sh
docker run -e CASUAL_PUBLIC_ORIGIN=https://sheets.acme.example ...
```

Or in the admin panel → **Networking** → _Public origin_. The
admin write wins over env.

## HSTS

Only emit when HTTPS terminates upstream — sending HSTS over HTTP
locks users out for the max-age window. Once you've verified
HTTPS works for two consecutive page loads, set
`CASUAL_HSTS_MAX_AGE=31536000` (1 year) in the admin panel
→ **Networking** → _HSTS max-age_.

Most modern proxies will emit HSTS themselves; check whether
yours does before setting it here too (duplicate headers are
harmless but slightly noisy).

## Sub-path vs sub-domain

Two shapes, both supported:

| Pattern | When | What changes |
|---|---|---|
| `https://sheets.acme.example/` | Default. Dedicated subdomain. | Nothing — defaults work. |
| `https://acme.example/sheets/` | Sharing a domain with other apps. | Set _Base path_ to `/sheets` in the admin panel. Configure the proxy to forward `/sheets/*` (and `/sheets/yjs`) verbatim, without stripping the prefix. |

The base-path mode also requires `CASUAL_PUBLIC_ORIGIN` to include
the prefix:

```
CASUAL_PUBLIC_ORIGIN=https://acme.example/sheets
```

## Verifying the cert chain

```sh
openssl s_client -showcerts -servername sheets.acme.example \
  -connect sheets.acme.example:443 < /dev/null \
  | openssl x509 -noout -subject -issuer -dates
```

Should print `notBefore` ≤ today ≤ `notAfter` with at least 30
days runway. All three proxies renew automatically; this is a
human-readable sanity check after first provisioning.
