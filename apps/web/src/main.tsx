import ReactDOM from 'react-dom/client';
import { App } from './App';
import { AdminApp } from './admin/AdminApp';
import { SdkHarness } from './sdk-harness/SdkHarness';
// Design-system tokens (Phase 4): loads Inter / JetBrains Mono / Manrope +
// Material Symbols + the 137 design tokens (colours, spacing, radius, shadows,
// motion, chrome heights) on :root. Imported BEFORE styles.css so the app +
// SDK chrome can reference `--color-*` / `--space-* / `--motion-*` etc., and
// `[data-theme="dark"]` swaps them.
import '@schnsrw/design-system/tokens.css';
import './styles.css';

// Pathname-based routing — mirrors the home/collab pattern in App.tsx.
// `/admin` (or `/admin/...`) mounts the admin panel; everything else
// renders the editor shell. One-shot at boot — moving between admin
// and editor reloads the page, which is fine because the admin panel
// is configuration, not a hot path.
const isAdminRoute = window.location.pathname.startsWith('/admin');
// Dev/test-only: mount the SDK's <CasualSheets> editor in isolation so Playwright
// can exercise the published component directly (the app otherwise renders its own
// UniverSheet). See apps/web/src/sdk-harness/SdkHarness.tsx.
const isSdkHarness = window.location.pathname === '/sdk-harness';

// Note: React.StrictMode is intentionally NOT used.
// Univer mounts its own internal React root inside the container we hand it.
// StrictMode's intentional double-invocation of effects in dev unmounts/remounts
// the Univer instance before its first render completes, leaving the DOM in an
// inconsistent state (insertBefore/removeChild on detached nodes). This is the
// same pattern most heavy editor SDKs (Monaco, Univer, Lexical with portals)
// document. Restore StrictMode only at child boundaries that don't host Univer.
ReactDOM.createRoot(document.getElementById('root')!).render(
  isSdkHarness ? <SdkHarness /> : isAdminRoute ? <AdminApp /> : <App />,
);
