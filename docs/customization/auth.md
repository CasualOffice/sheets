# Auth — JWT, roles, permissions, features

When `CASUAL_JWT_SECRET` is set, every WOPI request must carry a
signed JWT. Tokens encode the file binding + user identity +
role + permission flags + feature toggles + lifetime. The server
verifies + enforces; the client reads + drives UI gating.

When `CASUAL_JWT_SECRET` is **unset**, WOPI routes fall through
to v0.0.x anonymous-by-URL behaviour. Operators opt in to auth
by setting the secret.

## Claims model

A signed token carries:

| Claim | Type | Description |
|---|---|---|
| `sub` | string | Username, email, stable user id. Surfaces as WOPI `UserId`. |
| `file_id` | string | The single file this token authorises. WOPI routes reject when URL `:id` ≠ this claim. Use `*` only for admin tokens that need to mint other tokens. |
| `role` | enum | `admin` · `editor` · `commenter` · `viewer`. Drives the default permission map below. |
| `permissions` | object | Optional per-flag override. See "Roles" below. |
| `features` | object | Optional feature toggle override. See "Feature toggles" below. |
| `password_required` | boolean | When true, the legacy `x-room-password` header gate also applies on top of the JWT. |
| `display_name` | string | Human label for presence + cursor markers. Falls back to `sub`. |
| `aud` | string | Audience — usually the deployment's public origin. |
| `iss` | string | Issuer — useful when downstream SSO mints tokens. |
| `exp`, `iat` | number | Standard JWT expiry + issued-at, set by the signer. |

## Roles → default permissions

| Role | read | write | comment | download | share | admin |
|---|---|---|---|---|---|---|
| `admin` | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |
| `editor` | ✓ | ✓ | ✓ | ✓ | — | — |
| `commenter` | ✓ | — | ✓ | ✓ | — | — |
| `viewer` | ✓ | — | — | ✓ | — | — |
| `anonymous` | — | — | — | — | — | — |

The `permissions` claim overrides on a per-flag basis — useful
for tokens like "viewer who can't even download" (`viewer` +
`{ permissions: { download: false } }`).

## Feature toggles

| Feature | Drives |
|---|---|
| `charts` | Chart insert + format |
| `pivots` | Pivot insert + edit |
| `conditionalFormatting` | CF rules + DV rules |
| `sharing` | Share-link generation |
| `exportFiles` | Download as .xlsx / .ods / .csv |
| `collab` | Real-time co-edit for this session |
| `ai` | Inline AI features (when wired in v0.3+) |

Defaults (deployment-wide): everything `true` except `ai`. The
admin panel can flip defaults; per-token claims override
deployment defaults.

## Bootstrapping the first admin token

Tokens are minted by tokens — chicken-and-egg. Sign the first
admin token manually using the secret. Inside the container:

```sh
docker compose exec app node -e '
  import("jsonwebtoken").then(({ default: jwt }) => {
    const tok = jwt.sign(
      { sub: "owner", file_id: "*", role: "admin" },
      process.env.CASUAL_JWT_SECRET,
      { algorithm: "HS256", expiresIn: "8h" },
    );
    console.log(tok);
  });
'
```

(Note the `file_id: "*"` sentinel — admin tokens don't bind to a
single file. The check `ctx.fileIdMatches(":id")` is only enforced
on the WOPI routes, not on `/api/tokens` or other admin routes.)

After bootstrap, the admin panel at `/admin` mints subsequent
tokens through its UI. The bootstrap step happens once per
deployment.

## Minting a per-user token

```sh
curl -X POST http://localhost:3000/api/tokens \
  -H "Authorization: Bearer $ADMIN_TOK" \
  -H "content-type: application/json" \
  -d '{
    "sub": "alice@acme.example",
    "display_name": "Alice",
    "file_id": "wb-q3-budget",
    "role": "editor",
    "permissions": { "share": true },
    "features": {
      "ai": false,
      "exportFiles": true,
      "sharing": true
    },
    "ttl_seconds": 3600
  }'
```

Response:

```json
{
  "token": "eyJhbGciOiJIUzI1NiIs...",
  "ttl_seconds": 3600,
  "claims": {
    "sub": "alice@acme.example",
    "file_id": "wb-q3-budget",
    "role": "editor",
    "permissions": { "share": true },
    "features": { "ai": false, "exportFiles": true, "sharing": true },
    "display_name": "Alice",
    "iat": 1748872989
  },
  "resolved_permissions": {
    "read": true, "write": true, "comment": true,
    "download": true, "share": true, "admin": false
  },
  "resolved_features": {
    "charts": true, "pivots": true, "conditionalFormatting": true,
    "sharing": true, "exportFiles": true, "collab": true, "ai": false
  }
}
```

`resolved_permissions` + `resolved_features` show what the server
will enforce + the client will see — useful for confirming overrides
landed where you expected.

## Using a token

Two placements supported:

```sh
# Authorization header (preferred for programmatic clients)
curl http://localhost:3000/wopi/files/wb-q3-budget/contents \
  -H "Authorization: Bearer $TOK"

# Query string (WOPI standard, also useful for share links)
curl "http://localhost:3000/wopi/files/wb-q3-budget?access_token=$TOK"
```

The client's `/api/me` endpoint introspects the token + returns
resolved claims:

```sh
curl -H "Authorization: Bearer $TOK" http://localhost:3000/api/me
```

Returns:

```json
{
  "anonymous": false,
  "role": "editor",
  "sub": "alice@acme.example",
  "displayName": "Alice",
  "fileId": "wb-q3-budget",
  "permissions": { "read": true, "write": true, ... },
  "features": { "ai": false, "exportFiles": true, ... },
  "passwordRequired": false,
  "exp": 1748876589
}
```

The web app uses this to gate UI — disabled buttons for missing
permissions, hidden menu items for disabled features.

## CheckFileInfo response

The standard WOPI `CheckFileInfo` response includes the same
information for clients that follow the spec verbatim:

```json
{
  "BaseFileName": "Q3 Budget.xlsx",
  "Size": 12345,
  "Version": "1748872989-abc12345",
  "UserId": "alice@acme.example",
  "UserFriendlyName": "Alice",
  "ReadOnly": false,
  "UserCanWrite": true,
  "UserCanRename": false,
  "UserCanAttend": true,
  "UserCanPresent": false,

  "casualRole": "editor",
  "casualPermissions": { "read": true, "write": true, ... },
  "casualFeatures": { "charts": true, ... },
  "casualPasswordRequired": false
}
```

The `casual*` extensions are non-standard but parallel — clients
either follow the WOPI booleans or the Casual-specific
permission/feature objects.

## Common error responses

| Status | Error code | Meaning |
|---|---|---|
| 401 | `access token required` | No token in header / query. |
| 401 | `token verify failed: …` | Bad signature, expired, malformed. |
| 401 | `token verify failed: jwt expired` | Sign a fresh token. |
| 403 | `file_id_mismatch` | Token's `file_id` claim doesn't match URL `:id`. |
| 403 | `read_not_permitted` | Token's `permissions.read` is false. |
| 403 | `write_not_permitted` | Token's `permissions.write` is false (or role is `viewer` / `commenter` without override). |
| 403 | `admin_required` | Route requires admin role. |
| 409 | `version_mismatch` | `If-Match` (`X-WOPI-ItemVersion`) on PutFile didn't match the stored version. |

## Migration from anonymous → JWT

Casual Sheets supports a gradual migration:

1. **Phase 1 — JWT optional.** Don't set `CASUAL_JWT_SECRET`.
   Everything works as v0.0.x.
2. **Phase 2 — admin only.** Set `CASUAL_ADMIN_USERNAME` +
   `CASUAL_ADMIN_PASSWORD` + `CASUAL_JWT_SECRET`. Admin panel works.
   WOPI routes still allow anonymous (because the `*` route block
   doesn't require auth when the deployment hasn't opted into it
   for everyone). Actually no — once the secret is set, WOPI
   requires auth. So:
3. **Phase 2 (corrected) — everything authed.** Once the secret is
   set, all WOPI requests need a valid JWT. Mint per-user tokens
   via `/api/tokens` and pass them in your share-link issuer or
   reverse-proxy auth layer.
4. **Phase 3 (v0.2) — SSO.** The OIDC + SAML stubs in the admin
   panel will route incoming users through an identity provider +
   mint tokens automatically.
