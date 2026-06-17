/**
 * Bootstraps the @schnsrw/design-system tokens inside the embed runtime's
 * iframe context. We can't rely on tsup's `injectStyle` to bundle the
 * package's CSS (bare-specifier CSS imports aren't followed reliably), so
 * the tokens + font-faces are inlined here as a critical-CSS literal +
 * a Google Fonts `<link>` for the webfonts (Inter / JetBrains Mono /
 * Manrope / Material Symbols Outlined).
 *
 * Keep in sync with `@schnsrw/design-system@0.1.0` — the literal mirrors
 * the package's `tokens/{colors,typography,spacing,motion}.css` plus the
 * editor-theme.css cyan override for the docs editor. Bumping
 * design-system here means updating both this string AND the npm dep
 * once the package is published.
 */

const FONTS_HREF =
  'https://fonts.googleapis.com/css2?' +
  'family=Inter:wght@400;500;600;700&' +
  'family=JetBrains+Mono:wght@400;500;600;700&' +
  'family=Manrope:wght@600;700;800&' +
  'family=Material+Symbols+Outlined:opsz,wght,FILL,GRAD@20..48,100..700,0..1,-50..200&' +
  'display=swap';

const TOKENS_CSS = `
.material-symbols-outlined {
  font-family: 'Material Symbols Outlined';
  font-weight: normal;
  font-style: normal;
  font-size: 20px;
  line-height: 1;
  letter-spacing: normal;
  text-transform: none;
  display: inline-block;
  white-space: nowrap;
  word-wrap: normal;
  direction: ltr;
  -webkit-font-feature-settings: 'liga';
  -webkit-font-smoothing: antialiased;
  font-variation-settings: 'FILL' 0, 'wght' 400, 'GRAD' 0, 'opsz' 24;
}

:root {
  color-scheme: light;

  --color-accent: #0e7490;
  --color-accent-hover: #0c627a;
  --color-accent-active: #0a5266;
  --color-accent-fg: #ffffff;
  --color-accent-soft: #e6f3f7;
  --color-accent-bright: #1597ba;
  --accent-gradient: linear-gradient(135deg, #1597ba 0%, #0e7490 100%);

  --color-bg: #ffffff;
  --color-surface: #ffffff;
  --color-surface-alt: #f4f6f9;
  --color-surface-strip: #eef1f5;
  --color-surface-raised: #ffffff;

  --color-border: #e6e9ee;
  --color-border-strong: #cdd3db;
  --color-divider: #edeff3;

  --color-glass: rgba(255, 255, 255, 0.72);
  --color-glass-strong: rgba(255, 255, 255, 0.86);
  --color-glass-border: rgba(255, 255, 255, 0.6);
  --color-scrim: rgba(15, 23, 42, 0.38);

  --color-text: #201f1e;
  --color-text-secondary: #605e5c;
  --color-text-disabled: #8a8886;
  --color-text-muted: #8a8886;
  --color-text-on-accent: #ffffff;

  --color-hover: rgba(15, 23, 42, 0.045);
  --color-pressed: rgba(15, 23, 42, 0.09);
  --color-focus-ring: #0e7490;
  --color-selected: rgba(14, 116, 144, 0.11);
  --color-selected-strong: rgba(14, 116, 144, 0.20);
  --color-toolbar-pill: #eef1f5;

  --color-success: #15803d;
  --color-success-soft: #dcfce7;
  --color-warning: #b45309;
  --color-warning-soft: #fef3c7;
  --color-danger: #b91c1c;
  --color-danger-soft: #fee2e2;
  --color-info: #1d4ed8;
  --color-info-soft: #dbeafe;

  --suite-sheets: #0e7490;
  --suite-editor: #0891b2;
  --suite-slides: #b91c1c;
  --suite-desktop: #ea580c;

  --font-sans: 'Inter', 'Segoe UI', -apple-system, BlinkMacSystemFont, 'Helvetica Neue', Arial, sans-serif;
  --font-mono: 'JetBrains Mono', 'SFMono-Regular', Consolas, 'Liberation Mono', Menlo, monospace;
  --font-display: 'Manrope', 'Inter', 'Segoe UI', -apple-system, sans-serif;

  --text-xs: 11px;
  --text-sm: 12px;
  --text-base: 13px;
  --text-md: 14px;
  --text-lg: 16px;

  --display-sm: 28px;
  --display-md: 36px;
  --display-lg: 44px;

  --leading-tight: 1.2;
  --leading-normal: 1.4;
  --leading-relaxed: 1.6;

  --weight-regular: 400;
  --weight-medium: 500;
  --weight-semibold: 600;
  --weight-bold: 700;

  --tracking-tight: -0.022em;
  --tracking-normal: 0;
  --tracking-wide: 0.04em;

  --space-0: 0;
  --space-1: 2px;
  --space-2: 4px;
  --space-3: 6px;
  --space-4: 8px;
  --space-5: 12px;
  --space-6: 16px;
  --space-7: 20px;
  --space-8: 24px;
  --space-9: 32px;
  --space-10: 48px;

  --radius-sm: 4px;
  --radius-md: 6px;
  --radius-lg: 10px;
  --radius-xl: 14px;
  --radius-pill: 999px;

  --shadow-1: 0 1px 2px rgba(16, 24, 40, 0.05), 0 1px 1px rgba(16, 24, 40, 0.04);
  --shadow-2: 0 2px 4px rgba(16, 24, 40, 0.06), 0 6px 16px -4px rgba(16, 24, 40, 0.10);
  --shadow-3: 0 4px 8px rgba(16, 24, 40, 0.06), 0 16px 32px -8px rgba(16, 24, 40, 0.14);
  --shadow-4: 0 8px 16px rgba(16, 24, 40, 0.08), 0 28px 56px -12px rgba(16, 24, 40, 0.22);
  --glow-accent: 0 0 0 3px var(--color-accent-soft);

  --motion-fast: 90ms;
  --motion-base: 160ms;
  --motion-slow: 240ms;
  --ease-out: cubic-bezier(0.16, 1, 0.3, 1);
  --ease-in-out: cubic-bezier(0.45, 0, 0.2, 1);
  --ease-spring: cubic-bezier(0.34, 1.4, 0.5, 1);

  --titlebar-h: 64px;
  --toolbar-h: 66px;
  --formula-bar-h: 26px;
  --sheet-tabs-h: 28px;
  --statusbar-h: 24px;
}

[data-theme='dark'] {
  color-scheme: dark;

  --color-accent: #22d3ee;
  --color-accent-hover: #3ddff5;
  --color-accent-active: #1cb8d4;
  --color-accent-fg: #07181c;
  --color-accent-soft: rgba(34, 211, 238, 0.16);
  --color-accent-bright: #38e0f5;
  --accent-gradient: linear-gradient(135deg, #38e0f5 0%, #22d3ee 100%);

  --color-bg: #14161a;
  --color-surface: #1b1e23;
  --color-surface-alt: #23262c;
  --color-surface-strip: #2a2e35;
  --color-surface-raised: #24282e;
  --color-border: #32363d;
  --color-border-strong: #424751;
  --color-divider: #24272d;

  --color-glass: rgba(28, 32, 38, 0.66);
  --color-glass-strong: rgba(28, 32, 38, 0.85);
  --color-glass-border: rgba(255, 255, 255, 0.08);
  --color-scrim: rgba(5, 8, 12, 0.58);

  --color-text: #e6e6e6;
  --color-text-secondary: #b0b3ba;
  --color-text-disabled: #6b6e75;
  --color-text-muted: #8d9098;
  --color-text-on-accent: #0e0f12;

  --color-hover: rgba(255, 255, 255, 0.06);
  --color-pressed: rgba(255, 255, 255, 0.12);
  --color-focus-ring: #22d3ee;
  --color-selected: rgba(34, 211, 238, 0.20);
  --color-selected-strong: rgba(34, 211, 238, 0.34);
  --color-toolbar-pill: #2a2c31;

  --color-success: #4ade80;
  --color-success-soft: rgba(74, 222, 128, 0.16);
  --color-warning: #fbbf24;
  --color-warning-soft: rgba(251, 191, 36, 0.16);
  --color-danger: #f87171;
  --color-danger-soft: rgba(248, 113, 113, 0.16);
  --color-info: #60a5fa;
  --color-info-soft: rgba(96, 165, 250, 0.16);

  --shadow-1: 0 1px 2px rgba(0, 0, 0, 0.44), 0 1px 1px rgba(0, 0, 0, 0.30);
  --shadow-2: 0 2px 4px rgba(0, 0, 0, 0.50), 0 6px 16px -4px rgba(0, 0, 0, 0.55);
  --shadow-3: 0 4px 8px rgba(0, 0, 0, 0.50), 0 16px 32px -8px rgba(0, 0, 0, 0.62);
  --shadow-4: 0 8px 16px rgba(0, 0, 0, 0.55), 0 28px 56px -12px rgba(0, 0, 0, 0.70);
}

[data-app='docs'] {
  --color-accent: #0891b2;
  --color-accent-hover: #0779a0;
  --color-accent-active: #056384;
  --color-accent-fg: #ffffff;
  --color-accent-soft: #e6f6fa;
  --color-accent-bright: #22b8d8;
  --accent-gradient: linear-gradient(135deg, #22b8d8 0%, #0891b2 100%);

  --color-focus-ring: #0891b2;
  --color-selected: rgba(8, 145, 178, 0.11);
  --color-selected-strong: rgba(8, 145, 178, 0.20);
}

[data-app='docs'][data-theme='dark'] {
  --color-accent: #67e8f9;
  --color-accent-hover: #84efff;
  --color-accent-active: #54ccdd;
  --color-accent-fg: #052a32;
  --color-accent-soft: rgba(103, 232, 249, 0.16);
  --color-accent-bright: #84efff;
  --accent-gradient: linear-gradient(135deg, #84efff 0%, #67e8f9 100%);

  --color-focus-ring: #67e8f9;
  --color-selected: rgba(103, 232, 249, 0.20);
  --color-selected-strong: rgba(103, 232, 249, 0.34);
}

@keyframes cs-fade-in { from { opacity: 0; } to { opacity: 1; } }
@keyframes cs-pop-in { from { opacity: 0; transform: translateY(-4px) scale(0.97); } to { opacity: 1; transform: translateY(0) scale(1); } }
@keyframes cs-rise-in { from { opacity: 0; transform: translateY(10px) scale(0.985); } to { opacity: 1; transform: translateY(0) scale(1); } }
@keyframes cs-slide-in-right { from { opacity: 0; transform: translateX(14px); } to { opacity: 1; transform: translateX(0); } }
@keyframes cs-slide-up { from { opacity: 0; transform: translateY(12px); } to { opacity: 1; transform: translateY(0); } }
@keyframes cs-scrim-in { from { opacity: 0; } to { opacity: 1; } }
.cs-anim-fade { animation: cs-fade-in var(--motion-base) var(--ease-out) both; }
.cs-anim-pop { animation: cs-pop-in var(--motion-base) var(--ease-spring) both; transform-origin: top; }
.cs-anim-rise { animation: cs-rise-in var(--motion-slow) var(--ease-spring) both; }
.cs-anim-panel { animation: cs-slide-in-right var(--motion-slow) var(--ease-out) both; }
.cs-anim-up { animation: cs-slide-up var(--motion-base) var(--ease-out) both; }
.cs-anim-scrim { animation: cs-scrim-in var(--motion-fast) var(--ease-out) both; }

@media (prefers-reduced-motion: reduce) {
  .cs-anim-fade, .cs-anim-pop, .cs-anim-rise, .cs-anim-panel, .cs-anim-up, .cs-anim-scrim {
    animation: cs-fade-in 1ms linear both;
  }
}
`;

const MARKER_ATTR = 'data-design-system';

export function injectDesignSystemTokens(): void {
  if (typeof document === 'undefined') return;

  if (!document.head.querySelector(`link[${MARKER_ATTR}]`)) {
    const link = document.createElement('link');
    link.rel = 'stylesheet';
    link.href = FONTS_HREF;
    link.setAttribute(MARKER_ATTR, 'fonts');
    document.head.appendChild(link);
  }

  if (!document.head.querySelector(`style[${MARKER_ATTR}]`)) {
    const style = document.createElement('style');
    style.setAttribute(MARKER_ATTR, 'tokens');
    style.textContent = TOKENS_CSS;
    document.head.appendChild(style);
  }
}
