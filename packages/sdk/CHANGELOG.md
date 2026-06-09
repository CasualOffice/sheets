# @schnsrw/casual-sheets

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
