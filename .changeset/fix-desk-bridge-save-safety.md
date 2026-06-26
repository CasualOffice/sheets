---
'@sheet/web': patch
---

Harden the desktop save bridge against two data-loss vectors: reject an empty (0-byte) serialization in `chunkedWrite` instead of atomically committing it over the original file, and stop clearing the dirty flag after a save if an edit landed while the write was in flight (it would otherwise be marked saved and lost on window close).
