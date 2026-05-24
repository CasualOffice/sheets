# Customization overview

Casual Sheets ships with an admin panel at `/admin` for runtime
customization — branding, storage backend, networking, room
limits, auth providers, webhooks, base path. No restart needed
for most changes; the on-disk config is reloaded on every read.

## Enabling the admin panel

Three env vars to set:

```sh
CASUAL_ADMIN_USERNAME=admin
CASUAL_ADMIN_PASSWORD=$(openssl rand -hex 16)
CASUAL_JWT_SECRET=$(openssl rand -hex 32)
```

(The JWT secret is what the panel signs your session token with —
must be ≥ 16 chars; 32 random bytes is a reasonable production
floor.)

Then `https://your-host/admin` shows a login form. Credentials are
constant-time compared against the env values; on success, you get
a short-lived admin-role JWT (1 hour default, configurable via
`CASUAL_ADMIN_SESSION_TTL`).

If any of the three env vars are unset, `/admin` shows a
"not configured" hint listing exactly which env vars are missing.

## Section reference

The panel has seven sections. Each is documented in detail:

- **[Branding](./branding.md)** — app name, accent colour, logo.
- **[Storage](./storage.md)** — workbook persistence backend
  selection + per-backend creds.
- **[Networking](./networking.md)** — public origin, CORS, trust
  proxy, HSTS.
- **[Base path](./base-path.md)** — reverse-proxy sub-path mount.
- **[Room limits](./room-limits.md)** — max rooms / file size /
  TTL / users per room.
- **[Auth providers](./auth.md)** — JWT (live), OIDC + SAML
  (stubs for v0.2).
- **[Webhooks](./webhooks.md)** — event-driven HTTP POSTs with
  HMAC signing.

## How config layers

```
defaults (compiled into the image)
   ↓ overridden by
env vars (CASUAL_*)
   ↓ overridden by
admin-panel JSON config (CASUAL_ADMIN_CONFIG_PATH)
```

Operators bootstrap a deployment with env (good for config-as-code,
secrets-manager integration, etc.); the admin panel writes are the
runtime override. Setting an env var after the panel has overridden
it does **not** revert — the on-disk JSON wins until an admin
panel save overwrites it.

To force-reset the panel: delete the on-disk JSON file + restart.
The server re-creates it with defaults.

## Secret handling

The admin panel never re-displays secrets it's been given. Once
saved, the S3 secret key + OIDC client secret + webhook signing
secrets all return as `***` from `GET /api/admin/config`. The
panel sends `***` back verbatim on unchanged fields; the
patch endpoint detects the sentinel and preserves the prior
verbatim value.

To rotate a secret: type the new value over the `***`. To delete:
type an empty string + save.

## Issuing access tokens

Once `CASUAL_JWT_SECRET` is set, every `/wopi/files/*` request
requires a signed JWT. Admin-role tokens can mint subordinate
tokens via:

```sh
curl -X POST http://localhost:3000/api/tokens \
  -H "Authorization: Bearer $ADMIN_TOK" \
  -H "content-type: application/json" \
  -d '{
    "sub": "alice@acme.example",
    "file_id": "wb-q3-budget",
    "role": "editor",
    "features": { "ai": false, "exportFiles": true },
    "ttl_seconds": 3600
  }'
```

Returns:

```json
{
  "token": "eyJhbGciOiJIUzI1NiIs...",
  "ttl_seconds": 3600,
  "claims": { ... },
  "resolved_permissions": { "read": true, "write": true, ... },
  "resolved_features": { ... }
}
```

Drop the returned `token` into the WOPI request:

```sh
curl https://your-host/wopi/files/wb-q3-budget/contents \
  -H "Authorization: Bearer $TOK"
```

(Or as the standard WOPI placement, `?access_token=<JWT>` query
string — Casual Sheets accepts both.)

See [`auth.md`](./auth.md) for the full claim model + role
permissions matrix.

## Webhook signature verification

When you set a signing secret on a webhook subscription, every
dispatch carries `X-Casual-Signature: sha256=<hex>`. Verify with:

```js
// Node.js — Fastify route receiving Casual webhooks
import { createHmac, timingSafeEqual } from 'node:crypto';

function verify(req, raw, secret) {
  const sig = req.headers['x-casual-signature'];
  if (!sig?.startsWith('sha256=')) return false;
  const provided = sig.slice('sha256='.length);
  const expected = createHmac('sha256', secret).update(raw).digest('hex');
  // Constant-time compare — protects against timing oracles.
  return (
    provided.length === expected.length &&
    timingSafeEqual(Buffer.from(provided, 'hex'), Buffer.from(expected, 'hex'))
  );
}
```

```python
# Python — Flask route
import hmac, hashlib
def verify(req, raw, secret):
    sig = req.headers.get("X-Casual-Signature", "")
    if not sig.startswith("sha256="):
        return False
    provided = bytes.fromhex(sig[len("sha256="):])
    expected = hmac.new(secret.encode(), raw, hashlib.sha256).digest()
    return hmac.compare_digest(provided, expected)
```

The `raw` argument is the **raw request body bytes** (not the
parsed JSON object). Some web frameworks consume the stream before
your handler runs — make sure you're capturing the bytes before
the body parser does.

## Where to go next

- [`docs/ENV.md`](../ENV.md) — every env var, every default.
- [`docs/self-hosting/`](../self-hosting/) — running this in
  production: TLS, reverse proxy, scaling, backups.
- [`apps/server/src/admin/`](https://github.com/schnsrw/sheets/tree/main/apps/server/src/admin) —
  the source if you want to read what the panel actually does.
