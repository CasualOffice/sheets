---
'@sheet/web': patch
---

Close a narrow data-loss window in the desktop save path: the bridge pinned its "did the doc change" reference at `bridge.save()` entry, but the bytes are serialized earlier (chart render + xlsx encode happen in between). An edit landing in that gap could leave the on-disk file stale while the window was marked clean — and silently lost on close. The save caller now captures the edit counter at serialization time (`wb.save()`) and passes it through as `save(bytes, baselineSeq)`, so the bridge only clears dirty when nothing changed since the bytes were produced.
