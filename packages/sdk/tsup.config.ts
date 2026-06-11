import { defineConfig, type Plugin } from 'tsup';

/**
 * `parse-in-worker.ts` constructs the Web Worker via
 *
 *   new Worker(new URL('./parser.worker.ts', import.meta.url), { type: 'module' });
 *
 * which is canonical Vite syntax but breaks for any consumer that
 * ships the compiled .mjs (the .ts source isn't in node_modules).
 * Same fix as @schnsrw/docx-js-editor@1.0.1: emit the worker as a
 * sibling .mjs in dist (via the entry below) and rewrite the runtime
 * URL in the compiled chunk so the consumer's bundler resolves it.
 */
const rewriteParserWorkerUrl: Plugin = {
  name: 'rewrite-parser-worker-url',
  // The package.json has `"type": "module"` so tsup emits ESM with the
  // `.js` extension (and CJS with `.cjs`). The runtime URL points at
  // the ESM sibling since the Worker constructor with `type: 'module'`
  // demands ES-module syntax; CJS consumers that need a Worker have to
  // re-roll the construction. Same trade-off Vite makes for its own
  // worker plugin output.
  async renderChunk(code) {
    if (!code.includes('parser.worker.ts')) return null;
    const rewritten = code.replace(
      /["']\.\/parser\.worker\.ts["']/g,
      `'./parser.worker.js'`,
    );
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
    styles: 'src/styles.ts',
    xlsx: 'src/xlsx/index.ts',
    'parser.worker': 'src/xlsx/parser.worker.ts',
  },
  format: ['esm', 'cjs'],
  dts: true,
  splitting: false,
  sourcemap: true,
  clean: true,
  external: ['react', 'react-dom', /^@univerjs\//],
  plugins: [rewriteParserWorkerUrl],
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
  entry: { 'embed-runtime': 'src/embed-runtime/index.tsx' },
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
  // Empty external list — bundle every bare import into the runtime.
  external: [],
  noExternal: [/.*/],
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
  ],
});

export default [mainConfig, embedRuntimeConfig];
