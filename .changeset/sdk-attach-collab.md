---
'@casualoffice/sheets': minor
---

feat(collab): opt-in real-time co-editing via `@casualoffice/sheets/collab`

The editor ships collab-unaware. A host enables co-editing with one call after
`onReady`:

```ts
import { attachCollab } from '@casualoffice/sheets/collab';

const handle = attachCollab(api, { room: 'doc-42', server: 'wss://host/yjs' });
// …later
handle.detach();
```

- `attachCollab(api, { room, server, password?, role?, token?, onSnapshot?, onStatus? })`
  spins up the Yjs doc + Hocuspocus provider + mutation bridge and returns a
  `CollabHandle` (`doc`, `provider`, `bridge`, `status()`, `detach()`).
- The mutation bridge (`startBridge`) and replay machinery moved into the SDK —
  the non-negotiable Univer hooks (`onMutationExecutedForCollab`, `fromCollab`
  echo guard, `__splitChunk__`) travel with it.
- `yjs` and `@hocuspocus/provider` are **peer dependencies** (optional) so the
  host provides a single Yjs copy — two copies break `Y.Doc` identity.

Yjs/Hocuspocus is the realtime transport only; authoritative persistence stays
host-side (WOPI / backend) via the save/exit event contract.
