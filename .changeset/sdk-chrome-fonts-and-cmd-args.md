---
'@casualoffice/sheets': patch
---

Two embed/chrome fixes found integrating the SDK into a host (Drive):

- **Chrome font loader skipped Material Symbols.** `ensureChromeFonts` deduped on the bare `/css2` path, which is shared by both Google Fonts URLs, so the second family (Material Symbols Outlined) was never injected and `chrome="full"` icons rendered as raw ligature text. Now deduped per `family=` segment.
- **`CasualSheetsIframe` ref `executeCommand` dropped `args`.** It forwarded only `{ command }` over the postMessage protocol, so iframe-host commands carrying a payload (font family/size, colour) were no-ops. Now forwards `args` too.
