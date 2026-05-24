# Reverse-proxy recipes

Three proxies, three minimal configs. The WebSocket upgrade header
is the universal trip-wire — without it, the collab driver looks
"connected but silent" because HTTP-level traffic flows but the
upgrade-to-WebSocket leg never lands.

## What every proxy needs

1. **Forward the WebSocket upgrade** for `/yjs` (and `/yjs/*` for
   the docs paths).
2. **Bump the request body limit** above `MAX_UPLOAD_MB` (default
   100 MiB) for `POST /api/rooms/:id/seed` + `POST
   /wopi/files/:id/contents`.
3. **Preserve the X-Forwarded-\* headers** so the server knows the
   real client IP + scheme. Set `CASUAL_TRUST_PROXY` to match.
4. **Don't strip the base path** when running behind a sub-path
   (e.g. `https://acme/sheets`) — set the admin panel's
   _Base path_ to match and pass URLs through verbatim.
5. **Don't buffer the WebSocket stream**. Disable any
   `proxy_buffering` / `buffer` flags on the WS upstream block;
   buffering breaks the bidirectional sync protocol.

## nginx

```nginx
upstream casual_sheets_backend {
  server 127.0.0.1:3000;
  keepalive 32;
}

server {
  listen 443 ssl http2;
  server_name sheets.acme.example;

  ssl_certificate     /etc/letsencrypt/live/sheets.acme.example/fullchain.pem;
  ssl_certificate_key /etc/letsencrypt/live/sheets.acme.example/privkey.pem;

  # 100 MiB matches the default MAX_UPLOAD_MB.
  client_max_body_size 100m;

  # Trust headers from this proxy hop.
  set $real_scheme $scheme;

  location / {
    proxy_pass         http://casual_sheets_backend;
    proxy_http_version 1.1;

    # X-Forwarded-* set so the server can build correct redirect URLs.
    proxy_set_header   Host              $host;
    proxy_set_header   X-Real-IP         $remote_addr;
    proxy_set_header   X-Forwarded-For   $proxy_add_x_forwarded_for;
    proxy_set_header   X-Forwarded-Proto $real_scheme;
    proxy_set_header   X-Forwarded-Host  $host;

    # WebSocket upgrade — covers /yjs and any other upgrade
    # the app might add later.
    proxy_set_header   Upgrade $http_upgrade;
    proxy_set_header   Connection "upgrade";

    # Long-lived collab sockets need the read timeout bumped.
    proxy_read_timeout 86400s;
    proxy_send_timeout 86400s;

    # Disable buffering for WS upgrades. nginx is smart enough to
    # do this automatically when it sees Upgrade: websocket, but
    # the explicit setting documents the intent.
    proxy_buffering off;
  }
}
```

Then on the server: `CASUAL_TRUST_PROXY=loopback` (the default;
trust the nginx hop on 127.0.0.1).

### Sub-path mount

Serving at `https://acme.example/sheets` instead of a dedicated
subdomain:

```nginx
location /sheets/ {
  proxy_pass         http://casual_sheets_backend/sheets/;
  # ... (rest of the headers as above)
}
```

Note the **trailing slash on both sides**: `proxy_pass
http://upstream/sheets/` (not `/`) keeps the `/sheets` prefix in the
forwarded URL. Then set _Base path_ in the admin panel to `/sheets`.

## Caddy

```caddyfile
sheets.acme.example {
  encode gzip zstd

  # Bump from the default 10 MiB.
  request_body {
    max_size 100MB
  }

  reverse_proxy 127.0.0.1:3000 {
    # Caddy handles the WS upgrade transparently — no extra
    # directives needed.

    header_up X-Real-IP {remote_host}
    header_up X-Forwarded-For {remote_host}
    header_up X-Forwarded-Proto {scheme}
    header_up X-Forwarded-Host {host}

    # Long-lived collab sockets.
    transport http {
      read_timeout 24h
      write_timeout 24h
    }
  }
}
```

TLS auto-provisions from Let's Encrypt the first time the domain
is hit.

### Sub-path mount

```caddyfile
acme.example {
  handle_path /sheets/* {
    reverse_proxy 127.0.0.1:3000
  }
}
```

`handle_path` strips the prefix before forwarding. If you'd
rather keep the prefix (which the server expects with _Base path_
set), use `handle` instead and configure the server's base path
to match.

## Traefik

`docker-compose.yml` with the official labels:

```yaml
services:
  app:
    image: schnsrw/casual-sheets:0.1
    labels:
      - traefik.enable=true
      - traefik.http.routers.sheets.rule=Host(`sheets.acme.example`)
      - traefik.http.routers.sheets.tls=true
      - traefik.http.routers.sheets.tls.certresolver=letsencrypt
      - traefik.http.services.sheets.loadbalancer.server.port=3000
      # WebSocket upgrade — Traefik handles it for free, but the
      # body-size limit defaults to 1 MiB and needs a bump.
      - traefik.http.middlewares.body.buffering.maxRequestBodyBytes=104857600
      - traefik.http.routers.sheets.middlewares=body

  traefik:
    image: traefik:v3.1
    command:
      - --entrypoints.web.address=:80
      - --entrypoints.websecure.address=:443
      - --providers.docker=true
      - --certificatesresolvers.letsencrypt.acme.email=ops@acme.example
      - --certificatesresolvers.letsencrypt.acme.storage=/letsencrypt/acme.json
      - --certificatesresolvers.letsencrypt.acme.httpchallenge.entrypoint=web
    ports: ['80:80', '443:443']
    volumes:
      - /var/run/docker.sock:/var/run/docker.sock:ro
      - letsencrypt:/letsencrypt

volumes:
  letsencrypt:
```

WebSocket upgrade works out of the box with Traefik; no
additional config needed for `/yjs`.

## Checklist after wiring

1. `curl https://sheets.acme.example/health` → `{"ok": true, ...}`.
2. Open the app in two browsers. Edit a cell in one — peer cursor +
   live-typing ghost in the other within ~250 ms. (If silence, the
   WebSocket upgrade isn't landing.)
3. Upload a 50 MiB workbook via File → Open. (If the proxy returns
   413, bump `client_max_body_size` / `max_size` /
   `maxRequestBodyBytes` past 100 MiB.)
4. `curl -I https://sheets.acme.example/` shows
   `strict-transport-security` if you configured HSTS.
5. Admin panel at `/admin` loads + sign-in works.
