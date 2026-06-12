# @schnsrw/casual-sheets

## 0.5.4

### Patch Changes

- `noExternal: ['exceljs', /^@univerjs\//]` so the parser worker
  bundles @univerjs/core (it imports LocaleType + CustomRangeType).
  0.5.3 only added exceljs to noExternal; the worker still had
  `import { ... } from "@univerjs/core"` as a bare specifier and
  closed silently at load. The "OOM" error message the embed-runtime
  emitted was misleading — it was just an unresolvable bare import
  in a module-script worker.

## 0.5.3

### Patch Changes

- Three fixes to make the iframe embed actually render:
  1. `platform: 'browser'` on the main tsup config so the parser worker
     bundles exceljs's browser fork (no Node `stream` / `buffer` / `util`
     requires that broke worker init).
  2. embed-runtime calls `transport.sendReady()` after `sendHello()`. The
     host (CasualSheetsIframe) only sends its hello inside
     `onEditorReady`; without an eager `casual.ready` from the iframe,
     the handshake deadlocked and bytes never loaded.
  3. New tsup plugin copies `dist/parser.worker.js` into `dist/embed/`
     so the `new URL('./parser.worker.js', import.meta.url)` resolution
     inside the runtime finds the worker under `{embedBasePath}/`.

## 0.5.2

### Patch Changes

- Add `platform: 'browser'` to the embed-runtime tsup config so esbuild
  picks the browser variant of dual-target deps (nanoid, etc.). 0.5.1
  bundled everything but still grabbed `import { ... } from 'crypto'`
  from the Node fork of nanoid, which the browser can't resolve. The
  runtime now lands fully clean.

## 0.5.1

### Patch Changes

- Bundle React + Univer + all deps into the embed-runtime instead of
  leaving them as external imports. The previous build expected the
  consumer to provide an importmap; consumers like Casual Drive that
  embed via `<iframe src="…/embed.html">` had no way to do that, and
  the bare `import 'react'` failed at runtime in the browser.

  The runtime now ships ~11MB self-contained (cached after first load).

## 0.5.0

### Minor Changes

- e044efd: Ship the SDK iframe-delivery architecture for sheets (Phase 2 of doc 16
  in the parent docx repo). Mirror of `@schnsrw/docx-js-editor@1.1.0`.

  The existing `<CasualSheets>` direct-mount stays — no breaking change.
  Adds a new `<CasualSheetsIframe>` component that renders the editor
  inside a same-origin iframe. CSS isolation (Univer's design tokens
  stop bleeding into the host's tree), React-runtime isolation, and the
  Univer-vs-host font-cascade problems all go away when consumers
  switch from direct-mount to iframe.

  ### What the consumer-facing API looks like

  ```tsx
  import { CasualSheetsIframe } from '@schnsrw/casual-sheets';

  <CasualSheetsIframe
    fileSource={{
      open: async (id) => ({ bytes, name, etag }), // host's bytes shim
    }}
    docId={file.id}
    viewMode="preview"             // or "editor"
    embedBasePath="/embed/sheets"   // defaults to /embed/sheets
    onSelectionChanged={…}
    onError={…}
  />;
  ```

  No iframe, no postMessage, no `EmbedTransport` wiring in the consumer.
  Bytes flow host → iframe via `casual.load.request` envelopes; inside
  the iframe the runtime parses xlsx → `IWorkbookData` via the SDK's
  own `xlsxToWorkbookData` (Phase A of #56) and mounts `<CasualSheets>`
  with the snapshot.

  ### Build artifacts

  Two new files in `dist/embed/`:
  - `embed-runtime.js` (132 KB self-contained ESM) — mounts the editor
    inside the iframe; loads xlsx bytes via the wire, converts, renders.
  - `embed.html` — the 1 KB static HTML document the iframe loads.

  Consumers copy these into their public dir at `embedBasePath` (default
  `/embed/sheets`). A Vite plugin that does the copy ships in v0.5.x;
  for v0.5.0 the contract is a two-line postinstall:

  ```sh
  mkdir -p web/public/embed/sheets
  cp node_modules/@schnsrw/casual-sheets/dist/embed/* web/public/embed/sheets/
  ```

  ### Wire protocol additions
  - `casual.command.set.viewmode` — live preview ↔ editor toggle.
  - `casual.error` — editor → host fatal-error signal.

  Both mirror the docx repo's `13-iframe-protocol.md` extension.

  ### What's not in this minor
  - The full ref API (`flushSave`, `getSelection`, signing through iframe)
    — ships in v0.5.x once Drive proves the wire end-to-end.
  - The Vite plugin — v0.5.x.
  - Preview-mode chrome hiding inside the iframe — currently surfaced
    via `data-view-mode` attribute on the embed root + CSS gates in
    embed.html; v0.5.x wires the attribute to component-level `ui` props.
  - xlsx export from the iframe — Phase B of #56 still pending.

## 0.4.0

### Minor Changes

- Ships the xlsx **import** path as `@schnsrw/casual-sheets/xlsx` (Phase A of [#56](https://github.com/schnsrw/sheets/issues/56)).

  ```ts
  import { xlsxToWorkbookData } from '@schnsrw/casual-sheets/xlsx';

  const data = await xlsxToWorkbookData(arrayBuffer);
  // → IWorkbookData ready to mount via <CasualSheets initialData={data} />
  ```

  The parser runs in a Web Worker (`parser.worker.js`, bundled as a sibling
  in `dist/`). Consumer bundlers must support the
  `new Worker(new URL(...), import.meta.url)` pattern — Vite (with
  `worker.format: 'es'`), modern webpack with worker-plugin, esbuild's
  bundler.

  ### Fidelity scope
  - Values + formulas
  - Font (family, size, bold, italic, underline, colour)
  - Fill (solid background)
  - Alignment (horizontal, vertical, wrap)
  - Number format
  - Borders (thin, per side, colour preserved)
  - Merges
  - Sheet order + names
  - Tables, comments, data validation, page setup, named ranges (resources)

  Out of scope this release: charts, drawings, pivots, sparklines,
  advanced borders (dashed/double), themes, and **export** — Phase B of
  [#56](https://github.com/schnsrw/sheets/issues/56) handles export once
  the outline / charts / pivots / sparklines extension-point design is
  settled.

  ### What apps/web changed
  - `apps/web/src/xlsx/{import,parse-in-worker,parser.worker,parse-impl}.ts`
    and the shared utilities (`style-mapping`, `constants`, all 5
    `*-resource.ts` files, `pivot-passthrough.ts`) **moved** into
    `packages/sdk/src/xlsx/`.
  - `apps/web/src/xlsx/{export,export-impl}.ts` now imports the shared
    mappers + resource readers from `@schnsrw/casual-sheets/xlsx`. Same
    code, new path.
  - `apps/web/src/xlsx/index.ts` re-exports `xlsxToWorkbookData` from the
    SDK so existing apps/web call-sites are unaffected.

  ### Shared internals

  The SDK's `./xlsx` entry exports the shared style mappers + resource
  readers in addition to the importer. Hosts that ship their own xlsx
  export path (Casual Sheets' apps/web is one) use them to stay in
  lockstep with this importer's shape. Consumers that only need import
  ignore them — tree-shaking strips the unused symbols.

  ### Drive unblock

  [`schnsrw/drive`](https://github.com/schnsrw/drive) can now replace the
  `CasualSheetWorkspace` placeholder with a real loader:

  ```tsx
  const bytes = await driveFileSource.open(file.id);
  const data = await xlsxToWorkbookData(bytes);
  <CasualSheets initialData={data} ... />;
  ```

## 0.3.0

### Minor Changes

- 73e693f: Ships `CasualSheets` — a React wrapper around Univer Sheets. Mounts a
  workbook from `initialData`, boots the eager plugin set (render +
  formula engine + UI + docs + sheets + sheets-ui + sheets-formula +
  numfmt), and surfaces the `FUniver` API to the host via `onReady`.
  Hosts (Casual Drive in particular) can now `import { CasualSheets }
from '@schnsrw/casual-sheets/sheets'` and drop in a working
  spreadsheet view without re-implementing the boot dance.

  Lazy plugins (CF, drawings, sort, filter, hyperlinks, tables,
  comments, find/replace), the formula web worker, snapshot swap, and
  facade extensions stay app concerns — hosts layer them on top of
  `FUniver` after `onReady`.

  Also adds `./styles` (`import '@schnsrw/casual-sheets/styles'`) as a
  side-effect entry that brings in the eager plugin CSS in one line.

  Univer 0.24.x packages move to peer dependencies (all optional, all
  declared in `peerDependenciesMeta`).

## 0.2.0

### Minor Changes

- 06a5f3a: Initial release: `@schnsrw/casual-sheets` SDK shipping the signing pipeline
  (drawn / typed / uploaded signature surfaces, sequential / concurrent modes)
  and the iframe postMessage protocol (`EmbedTransport`, `casual.*` envelope
  types). Wire shapes are byte-identical to `@schnsrw/docx-js-editor` — only
  the `app` discriminator (`'sheet'` vs `'docs'`) and signature anchor shape
  (`{ kind: 'sheet', sheet, cell }` vs `{ kind: 'doc', paraId }`) differ. The
  Univer-Sheets React wrapper (`CasualSheets` component) is planned for a
  follow-up release.
