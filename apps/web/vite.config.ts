import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';

// `PAGES_BASE` lets the GitHub Pages workflow build for /sheets/ without
// committing that path into the repo (local dev stays at /).
const base = process.env.PAGES_BASE ?? '/';

// Read app version from package.json so the About dialog stays in sync
// without manual bumps every release.
const pkg = JSON.parse(
  readFileSync(
    resolve(dirname(fileURLToPath(import.meta.url)), 'package.json'),
    'utf-8',
  ),
) as { version: string };

const collabEnabled =
  process.env.VITE_COLLAB_ENABLED === '1' ||
  process.env.VITE_COLLAB_ENABLED === 'true';

export default defineConfig({
  base,
  plugins: [react()],
  define: {
    __APP_VERSION__: JSON.stringify(pkg.version),
    __COLLAB_BUILD__: JSON.stringify(collabEnabled),
  },
  server: {
    host: '127.0.0.1',
    port: 5273,
    strictPort: true,
  },
  // Formula offload worker is loaded with `{ type: 'module' }`, so the
  // worker bundle must be ES (not the default IIFE). IIFE can't code-split
  // and rollup hard-errors when an ES worker imports anything multi-chunk.
  worker: {
    format: 'es',
  },
  // Pre-bundle the heavy / dynamically-loaded deps at server start instead
  // of letting Vite discover them mid-run. Without this, the first dynamic
  // import of '@e965/xlsx' (when the user picks an .ods / .csv file or our
  // e2e suite probes the ods module) triggers a re-optimize pass that
  // duplicates Univer modules in the dep cache — the symptom is "Identifier
  // ... already exists" DI errors and a blank grid until a hard reload.
  optimizeDeps: {
    // Same rationale as @e965/xlsx — pre-bundle echarts at dev-server
    // start. Without this, the first dynamic import of echarts (when
    // a user clicks Insert > Chart, or when a workbook with charts
    // mounts) triggers a re-optimize pass mid-run that duplicates
    // Univer modules in the dep cache. Symptom: "Identifier ...
    // already exists" DI errors and a blank grid until hard reload.
    include: ['@e965/xlsx', 'echarts', 'echarts/core', 'echarts/charts', 'echarts/components', 'echarts/renderers'],
  },
});
