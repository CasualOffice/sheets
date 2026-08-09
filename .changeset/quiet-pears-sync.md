---
'@casualoffice/sheets': minor
---

Replace positional collaboration replay with a versioned, stable-id protocol that awaits mutations in order, rebuilds on concurrent insertion reorder, uses a private replay token for echo suppression, and fails closed at unknown or failed records. Compaction now uses preservation-aware snapshots, exposes auto/manual/off policy, defaults to off, and refuses unsafe browser state. Existing v1 rooms require a fresh room after upgrade.
