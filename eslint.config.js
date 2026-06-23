// Flat config (ESLint 9). Single source of truth for the whole workspace.
import js from '@eslint/js';
import globals from 'globals';
import tseslint from 'typescript-eslint';
import reactHooks from 'eslint-plugin-react-hooks';
import reactRefresh from 'eslint-plugin-react-refresh';
import prettier from 'eslint-config-prettier';

export default tseslint.config(
  {
    ignores: [
      '**/node_modules/**',
      '**/dist/**',
      '**/build/**',
      '**/.turbo/**',
      'vendor/**',
      // The collab server (apps/server) is a git submodule — CasualOffice/collab,
      // a product-agnostic server shared with Docs. It owns its own lint/format/CI;
      // we don't restyle vendored code. (Same treatment as vendor/**.)
      'apps/server/**',
      // Agent git worktrees live here — they're full copies of the repo (incl.
      // built JS), so linting them double-counts everything and breaks `lint`.
      '.claude/**',
      // Standalone embed demo: not a workspace member, own tsconfig, validated
      // separately (examples/embed-playground/playwright.embed.config.ts).
      'examples/**',
      'test-results/**',
      'playwright-report/**',
      '**/*.config.js',
      '**/*.config.ts',
    ],
  },
  js.configs.recommended,
  ...tseslint.configs.recommended,
  {
    files: ['**/*.{ts,tsx}'],
    languageOptions: {
      ecmaVersion: 2022,
      sourceType: 'module',
      globals: { ...globals.browser, ...globals.node },
    },
    plugins: {
      'react-hooks': reactHooks,
      'react-refresh': reactRefresh,
    },
    rules: {
      ...reactHooks.configs.recommended.rules,
      'react-refresh/only-export-components': ['warn', { allowConstantExport: true }],
      '@typescript-eslint/consistent-type-imports': ['warn', { prefer: 'type-imports' }],
      '@typescript-eslint/no-unused-vars': ['warn', { argsIgnorePattern: '^_', varsIgnorePattern: '^_' }],
    },
  },
  // Build / CI scripts run in Node — they need the Node globals (process,
  // console, URL, …) which the browser-leaning block above doesn't pick up
  // for plain .mjs files.
  {
    files: ['scripts/**/*.{js,mjs,cjs}'],
    languageOptions: {
      ecmaVersion: 2022,
      sourceType: 'module',
      globals: { ...globals.node },
    },
  },
  prettier,
);
