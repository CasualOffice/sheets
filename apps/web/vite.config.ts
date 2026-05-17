import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

// `PAGES_BASE` lets the GitHub Pages workflow build for /sheets/ without
// committing that path into the repo (local dev stays at /).
const base = process.env.PAGES_BASE ?? '/';

export default defineConfig({
  base,
  plugins: [react()],
  server: {
    host: '127.0.0.1',
    port: 5273,
    strictPort: true,
  },
  // Pre-bundle the heavy / dynamically-loaded deps at server start instead
  // of letting Vite discover them mid-run. Without this, the first dynamic
  // import of '@e965/xlsx' (when the user picks an .ods / .csv file or our
  // e2e suite probes the ods module) triggers a re-optimize pass that
  // duplicates Univer modules in the dep cache — the symptom is "Identifier
  // ... already exists" DI errors and a blank grid until a hard reload.
  optimizeDeps: {
    include: ['@e965/xlsx'],
  },
});
