import { defineConfig, type Plugin } from 'tsup';

/**
 * `parse-in-worker.ts` constructs the Web Worker via
 *
 *   new Worker(new URL('./parser.worker.ts', import.meta.url), { type: 'module' });
 *
 * which is canonical Vite syntax but breaks for any consumer that
 * ships the compiled .mjs (the .ts source isn't in node_modules).
 * Same fix as @casualoffice/docs@1.0.1: emit the worker as a
 * sibling .mjs in dist (via the entry below) and rewrite the runtime
 * URL in the compiled chunk so the consumer's bundler resolves it.
 */
const rewriteParserWorkerUrl: Plugin = {
  name: 'rewrite-worker-urls',
  // The package.json has `"type": "module"` so tsup emits ESM with the
  // `.js` extension (and CJS with `.cjs`). The runtime URL points at
  // the ESM sibling since the Worker constructor with `type: 'module'`
  // demands ES-module syntax; CJS consumers that need a Worker have to
  // re-roll the construction. Same trade-off Vite makes for its own
  // worker plugin output.
  async renderChunk(code) {
    if (!code.includes('parser.worker.ts') && !code.includes('exporter.worker.ts')) return null;
    const rewritten = code
      .replace(/["']\.\/parser\.worker\.ts["']/g, `'./parser.worker.js'`)
      .replace(/["']\.\/exporter\.worker\.ts["']/g, `'./exporter.worker.js'`);
    return { code: rewritten };
  },
};

import { promises as fs } from 'node:fs';

// Main library entries — code-split as before. The embed-runtime ships
// from a second config below so consumers get a self-contained ESM
// blob to drop into their public dir (not 100+ shared chunks).
const mainConfig = defineConfig({
  entry: {
    index: 'src/index.ts',
    signing: 'src/signing/index.ts',
    embed: 'src/embed/index.ts',
    sheets: 'src/sheets/index.ts',
    chrome: 'src/chrome/index.ts',
    collab: 'src/collab/index.ts',
    styles: 'src/styles.ts',
    xlsx: 'src/xlsx/index.ts',
  },
  format: ['esm', 'cjs'],
  dts: true,
  splitting: false,
  sourcemap: true,
  // Cleaning is done once via the `build` script (rmSync dist) before tsup runs.
  // Per-config clean:true here would race the parallel worker/univer/embed
  // configs and wipe their freshly-emitted dist files (e.g. univer.d.ts).
  clean: false,
  // Library entries (index/signing/embed/sheets/xlsx) MUST externalise @univerjs
  // (+ react). These are imported by host apps that already ship their own
  // @univerjs (the fork). Bundling Univer here would put a SECOND redi/@univerjs
  // copy in the host graph → "[redi]: loading scripts of redi more than once" +
  // LocaleService failures (rendering CasualSheets in a Univer-having host broke
  // exactly this way). The host resolves @univerjs to its single copy. exceljs is
  // a real `dependencies` entry — externalised by default so the consumer
  // resolves one copy too. Only the workers (separate configs below) bundle these
  // because a module worker has no import map at runtime.
  // `yjs` + `@hocuspocus/provider` are externalised for the same reason as
  // @univerjs: the `collab` entry is a peer-provided realtime layer. Two copies
  // of Yjs in the graph break `Y.Doc` identity (cross-copy `instanceof` fails,
  // awareness silently desyncs), so the host MUST resolve a single copy.
  // `@casualoffice/sheets/chrome` is externalised so `sheets`'s lazy
  // `import('@casualoffice/sheets/chrome')` is NOT inlined (this config is
  // splitting:false, which inlines relative dynamic imports) — it stays a bare
  // subpath import the CONSUMER's bundler code-splits, so `chrome="none"` hosts
  // never load the chrome chunk. Only this exact subpath is externalised; the
  // loader/other internals stay relative (bundled) as before.
  external: [
    'react',
    'react-dom',
    /^@univerjs\//,
    'yjs',
    '@hocuspocus/provider',
    '@casualoffice/sheets/chrome',
    // Same reason as /chrome: `sheets`'s `api.importXlsx` / `api.exportXlsx`
    // lazy `import('@casualoffice/sheets/xlsx')`. Externalising keeps it a bare
    // subpath the CONSUMER code-splits (splitting:false would otherwise inline
    // ~200KB of ExcelJS into the editor entry for hosts that never touch a file).
    '@casualoffice/sheets/xlsx',
  ],
  platform: 'browser',
  target: 'es2020',
  // Rewrites the './parser.worker.ts' URL in the xlsx code to the built '.js'
  // sibling emitted by workerConfig.
  plugins: [rewriteParserWorkerUrl],
});

// Workers — bundled, separate from the library entries. A module worker has no
// import map at runtime (iframe / Worker context), so its imports MUST be inlined.
// This is the one place @univerjs + exceljs are bundled; the library entries
// above keep them external to avoid duplicating redi in the host.
const workerConfig = defineConfig({
  entry: {
    'parser.worker': 'src/xlsx/parser.worker.ts',
    'exporter.worker': 'src/xlsx/exporter.worker.ts',
  },
  format: ['esm', 'cjs'],
  dts: false,
  splitting: false,
  sourcemap: true,
  clean: false,
  external: ['react', 'react-dom'],
  noExternal: ['exceljs', /^@univerjs\//],
  platform: 'browser',
  target: 'es2020',
});

// `./univer` — the shared Univer wiring (lazy plugin loader, Phase 1 Batch 1).
// MUST externalise @univerjs/react: unlike the parser worker (which bundles
// @univerjs/core because the iframe has no module map), this entry is imported
// by a host app that already ships its own @univerjs from the fork. Bundling
// Univer here would put a SECOND copy in the host's graph and break redi DI with
// duplicate-Identifier (Service2-suffix) errors. So everything @univerjs stays
// external and resolves to the host's single copy at runtime — including the
// lazy `import('@univerjs/...')` calls, which the host's bundler code-splits.
const univerLibConfig = defineConfig({
  entry: { univer: 'src/univer/index.ts' },
  format: ['esm', 'cjs'],
  dts: true,
  splitting: false,
  sourcemap: true,
  clean: false,
  external: ['react', 'react-dom', /^@univerjs\//],
  platform: 'browser',
  target: 'es2020',
});

// Embed-runtime — the in-iframe entry. Doc 16 §6. One self-contained
// file the consumer copies into `{embedBasePath}/embed-runtime.js`
// alongside `embed.html` (also copied below).
//
// Bundle EVERYTHING (react, react-dom, Univer) into the runtime — the
// previous build externalised these expecting the consumer to provide
// them via importmap. That broke browser loads because consumers like
// drive (which only does `<iframe src="…/embed.html">`) have no
// importmap; the bare `import 'react'` fails at runtime. The iframe
// is its own runtime context — bundling react / Univer there does NOT
// conflict with the host's copy. Trade-off: the runtime grows to
// ~10MB+, downloaded once per iframe load (cached after).
const embedRuntimeConfig = defineConfig({
  // Emit BOTH the embed runtime AND the parser worker into dist/embed/.
  // mainConfig also emits parser.worker.js to dist/, but tsup runs the
  // two configs in parallel — the embed config's `buildEnd` fires
  // before mainConfig finishes, so any "copy after" plugin races with
  // mainConfig's emission and silently fails. Emitting from the embed
  // config directly removes the race.
  entry: {
    'embed-runtime': 'src/embed-runtime/index.tsx',
    'parser.worker': 'src/xlsx/parser.worker.ts',
    'exporter.worker': 'src/xlsx/exporter.worker.ts',
  },
  outDir: 'dist/embed',
  format: ['esm'],
  // No dts — the runtime is a side-effect script loaded by embed.html
  // via <script type="module">, not imported by user code. Skipping
  // dts also dodges a tsup parallel-clean race that wipes
  // dist/embed/*.d.ts after the main config's clean=true fires.
  dts: false,
  splitting: false,
  sourcemap: false,
  clean: false,
  minify: true,
  // Bundle every bare import into both the runtime and the worker so
  // the in-iframe ES module loader doesn't try to resolve specifiers
  // it has no map for. Same posture as the parser worker emitted by
  // mainConfig — but emitted from THIS config so the file actually
  // lands in dist/embed/.
  external: [],
  noExternal: [/.*/],
  // Browser target: pick the `browser` field from each dep's package.json
  // so packages with Node/browser splits (nanoid, etc.) load the
  // browser variant. Without this, esbuild grabs `import { ... } from
  // 'crypto'` from the Node fork of nanoid and the iframe load fails.
  platform: 'browser',
  target: 'es2020',
  // Inline Univer's CSS modules into the runtime bundle. Without this
  // the embed.html has no <link rel="stylesheet"> and Univer's
  // workbench renders unstyled (the canvas mounts at 0x0 size, the
  // chrome divs have no visible layout). With injectStyle, every
  // imported CSS module turns into a JS string that gets appended to
  // a <style> tag at runtime — fully self-contained.
  injectStyle: true,
  loader: { '.css': 'css' },
  plugins: [
    rewriteParserWorkerUrl,
    {
      name: 'copy-embed-html',
      async buildEnd() {
        try {
          await fs.mkdir('dist/embed', { recursive: true });
          await fs.copyFile('src/embed-runtime/embed.html', 'dist/embed/embed.html');
        } catch (err) {
          if ((err as NodeJS.ErrnoException)?.code !== 'ENOENT') throw err;
        }
      },
    },
    {
      // The embed-runtime spawns the xlsx parser via
      // `new Worker(new URL('./parser.worker.js', import.meta.url))`,
      // which resolves to `<embedBasePath>/parser.worker.js`. The
      // mainConfig emits parser.worker to `dist/`, but consumers only
      // copy the `dist/embed/` tree to their public path — so the
      // worker URL 404s at runtime. Copy the worker into
      // `dist/embed/` so the relative path resolves.
      name: 'copy-parser-worker-into-embed',
      async buildEnd() {
        try {
          await fs.mkdir('dist/embed', { recursive: true });
          await fs.copyFile('dist/parser.worker.js', 'dist/embed/parser.worker.js');
        } catch (err) {
          if ((err as NodeJS.ErrnoException)?.code !== 'ENOENT') throw err;
        }
      },
    },
  ],
});

export default [mainConfig, workerConfig, univerLibConfig, embedRuntimeConfig];
