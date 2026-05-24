# CORS

CORS only matters when the **browser** is making cross-origin
requests to Casual Sheets from a different origin than the web app
is served from. If you're using the SPA + the API on the same
domain (the default shape), you don't need any CORS config — same-
origin requests don't trigger the policy.

Three scenarios where it matters:

## 1. Embedded editor

You're hosting Casual Sheets at `https://sheets.acme.example` and
embedding it as an `<iframe>` inside an app served from
`https://app.acme.example`. The iframe itself isn't CORS-restricted
for navigation, but any `fetch()` from the iframe code to
sibling apps' APIs is. The configuration here only matters if **the
iframe** makes API calls back to its own origin — which is the
default case and is _not_ cross-origin.

Result: no CORS needed.

## 2. External API consumer

You're calling the Casual Sheets API from a different origin —
e.g. an admin dashboard at `https://ops.acme.example` calling
`/api/files` on `https://sheets.acme.example` to list workbooks.

Set:

```
CASUAL_CORS_ORIGINS=https://ops.acme.example
```

(or via the admin panel → **Networking** → _CORS origins_). The
server then sends `Access-Control-Allow-Origin: https://ops.acme.example`
on matching requests + handles preflight (`OPTIONS`) automatically.

Multiple origins are comma-separated:

```
CASUAL_CORS_ORIGINS=https://ops.acme.example,https://staging.acme.example
```

## 3. Local dev against a remote server

You're running the web SPA locally on `http://localhost:5173`
(`pnpm dev:web`) against a remote staging server. Add the localhost
origin:

```
CASUAL_CORS_ORIGINS=http://localhost:5173
```

## Wildcard caveat

The current implementation does **not** support `*` as a value —
that's intentional. Any deployment that needs `*` is probably better
served by a same-origin shape (move the consumer app under the
same domain) because `*` disables credential forwarding and
makes XSRF harder to reason about.

## Most common mistake

> "I set CORS but my admin panel still fails with `Failed to fetch`"

That's not a CORS error. CORS errors say `blocked by CORS policy`
explicitly in the console. `Failed to fetch` is usually:

- DNS doesn't resolve
- The server isn't reachable from your browser's network
- A mixed-content block (HTTP API from an HTTPS page)
- Certificate is invalid (browser refuses the TLS handshake)

Test with `curl` from the same machine — if `curl` works, the
network is fine and your error is somewhere else.

## Preflight cache

`Access-Control-Max-Age` defaults to 86400 (24 h) — preflights
hit the server once per day per origin per request shape. If you
need to debug preflight behaviour, append a `cache-bust` query
param or use the Chrome devtools "Disable cache" toggle.
