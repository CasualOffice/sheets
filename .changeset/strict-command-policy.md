---
'@casualoffice/sheets': minor
---

Enforce disabled sheet feature flags atomically at Univer's shared command boundary, guard replay provenance, add a synchronous `onBeforeCommand` host veto, and expose a persistence-safe `onLocalMutation` stream without changing `onMutation` audit behavior.
