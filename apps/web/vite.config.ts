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
});
