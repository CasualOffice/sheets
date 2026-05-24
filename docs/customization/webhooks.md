# Webhooks

Server-side events fire HTTP POSTs to operator-configured URLs.
HMAC-SHA256 signs the JSON body when a subscription has a secret;
receivers verify via `X-Casual-Signature: sha256=<hex>`.

Subscriptions are managed in the admin panel → **Webhooks**. Each
subscription has:

| Field | Description |
|---|---|
| `name` | Human label for the subscription. Surfaces in the panel listing + server logs. |
| `url` | The endpoint Casual POSTs JSON to. |
| `events` | Array of event names. Empty = subscribed to every event. |
| `secret` | Optional HMAC-SHA256 signing secret. Strongly recommended. |
| `enabled` | Master switch. Disable to pause without removing the row. |

## Events

| Event | Fired when |
|---|---|
| `room.created` | `POST /api/rooms` creates a new room |
| `room.dropped` | Last client leaves + GC ticks |
| `file.uploaded` | `POST /api/rooms/:id/seed` succeeds |
| `file.saved` | `POST /wopi/files/:id/contents` succeeds |
| `file.deleted` | Admin deletes a file via the API |
| `user.joined` | New client joins a room |
| `user.left` | Client disconnects from a room |
| `admin.login` | Successful `/api/admin/login` |
| `admin.login_failed` | Failed `/api/admin/login` |

## Payload shape

```json
{
  "event": "file.saved",
  "timestamp": "2026-06-01T14:23:09.123Z",
  "payload": {
    "fileId": "wb-q3-budget",
    "size": 12345,
    "version": "1748872989-abc12345",
    "user": "alice@acme.example"
  }
}
```

The `payload` shape is event-specific. Stable across patch versions;
new fields may be added (your receiver should ignore unknown keys).

## Headers on every dispatch

| Header | Value |
|---|---|
| `Content-Type` | `application/json` |
| `User-Agent` | `CasualSheets-Webhook/0.1` |
| `X-Casual-Event` | The event name |
| `X-Casual-Attempt` | `1` (first try) or `2` (retry) |
| `X-Casual-Signature` | `sha256=<hex>` — only present when the subscription has a secret |

## Signature verification

The signature is `hmac-sha256(secret, raw_body)`. Verify it
constant-time against the header. Examples:

### Node.js (Fastify)

```js
import { createHmac, timingSafeEqual } from 'node:crypto';
import Fastify from 'fastify';

const app = Fastify();

// Capture the raw body bytes — content-type parsing reads the stream.
app.addContentTypeParser(
  'application/json',
  { parseAs: 'string' },
  (_req, raw, done) => {
    try {
      done(null, { raw, parsed: JSON.parse(raw) });
    } catch (err) {
      done(err);
    }
  },
);

app.post('/hook', async (req, reply) => {
  const sig = req.headers['x-casual-signature'];
  const { raw, parsed } = req.body;
  if (!sig?.startsWith('sha256=')) return reply.code(400).send('no sig');
  const provided = Buffer.from(sig.slice('sha256='.length), 'hex');
  const expected = createHmac('sha256', process.env.CASUAL_HOOK_SECRET)
    .update(raw)
    .digest();
  if (provided.length !== expected.length || !timingSafeEqual(provided, expected)) {
    return reply.code(401).send('bad sig');
  }
  console.log('verified webhook:', parsed.event, parsed.payload);
  return { ok: true };
});

app.listen({ port: 4000 });
```

### Python (Flask)

```python
import hmac, hashlib, os
from flask import Flask, request, abort

app = Flask(__name__)
SECRET = os.environ["CASUAL_HOOK_SECRET"].encode()

@app.route("/hook", methods=["POST"])
def hook():
    sig = request.headers.get("X-Casual-Signature", "")
    if not sig.startswith("sha256="):
        abort(400)
    provided = bytes.fromhex(sig[len("sha256="):])
    raw = request.get_data()                  # raw bytes, BEFORE json parsing
    expected = hmac.new(SECRET, raw, hashlib.sha256).digest()
    if not hmac.compare_digest(provided, expected):
        abort(401)
    body = request.get_json()
    print("verified webhook:", body["event"], body["payload"])
    return {"ok": True}
```

### Go (net/http)

```go
package main

import (
  "crypto/hmac"
  "crypto/sha256"
  "encoding/hex"
  "io"
  "net/http"
  "os"
  "strings"
)

func main() {
  secret := []byte(os.Getenv("CASUAL_HOOK_SECRET"))
  http.HandleFunc("/hook", func(w http.ResponseWriter, r *http.Request) {
    sig := r.Header.Get("X-Casual-Signature")
    if !strings.HasPrefix(sig, "sha256=") {
      http.Error(w, "no sig", http.StatusBadRequest)
      return
    }
    provided, err := hex.DecodeString(sig[len("sha256="):])
    if err != nil {
      http.Error(w, "bad sig encoding", http.StatusBadRequest)
      return
    }
    raw, _ := io.ReadAll(r.Body)
    mac := hmac.New(sha256.New, secret)
    mac.Write(raw)
    expected := mac.Sum(nil)
    if !hmac.Equal(provided, expected) {
      http.Error(w, "bad sig", http.StatusUnauthorized)
      return
    }
    w.WriteHeader(200)
    w.Write([]byte(`{"ok":true}`))
  })
  http.ListenAndServe(":4000", nil)
}
```

**Critical**: in every language, the HMAC input is the **raw body
bytes**, not the parsed object. Frameworks that JSON-parse before
your handler runs need explicit raw-capture (see the Node example
above).

## Retry policy

- Receiver returns 2xx within 8 s → dispatch logged + done.
- Non-2xx OR network error OR timeout → single retry after 5 s.
- Retry fails too → logged + dropped (no further retries in v0.1).

v0.2 ships a proper queue with exponential back-off + dead-letter
storage for receivers that go down longer than a few seconds.

## Operational considerations

- **Don't make the webhook handler slow** — the dispatcher's
  per-request timeout is 8 s. A slow receiver doesn't block the
  Casual request that triggered the event (dispatches run in a
  separate fire-and-forget task), but slow handlers do back up the
  Node event loop on the receiver side.
- **Idempotency** — events are at-least-once. The same event may
  fire twice if a retry succeeds after the first dispatch already
  reached the receiver but the response got lost. Use the
  `timestamp` + `payload.fileId` (or other identifying field) as
  an idempotency key on your side.
- **Order is not guaranteed.** Two events from the same room may
  arrive at the receiver out of order. Use the `timestamp` field
  for ordering decisions.

## Most common mistakes

1. **JSON-parsing the body before computing the HMAC.** The hash
   has to be on the raw bytes, not the re-serialised JSON
   (whitespace + key ordering changes invalidate the hash).
2. **Forgetting `enabled: true`.** Disabled subscriptions don't
   fire — easy to forget when wiring up a new one.
3. **HTTPS endpoint with a self-signed cert.** Node's `fetch`
   (which the dispatcher uses) rejects untrusted certs. Either
   serve the receiver via a real CA or work over HTTP on a
   private network.
