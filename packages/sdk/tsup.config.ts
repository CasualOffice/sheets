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

export default defineConfig({
  entry: {
    index: 'src/index.ts',
    signing: 'src/signing/index.ts',
    embed: 'src/embed/index.ts',
    sheets: 'src/sheets/index.ts',
    styles: 'src/styles.ts',
    xlsx: 'src/xlsx/index.ts',
    // Worker entry — emits dist/parser.worker.mjs + .cjs sibling.
    // The runtime URL in parse-in-worker.ts is rewritten via
    // rewriteParserWorkerUrl above so consumer bundlers can resolve it.
    'parser.worker': 'src/xlsx/parser.worker.ts',
  },
  format: ['esm', 'cjs'],
  dts: true,
  splitting: false,
  sourcemap: true,
  clean: true,
  external: [
    'react',
    'react-dom',
    // Univer is peer; consumers install the matching @univerjs/* set.
    /^@univerjs\//,
  ],
  plugins: [rewriteParserWorkerUrl],
});
