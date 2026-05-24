#!/usr/bin/env node
/**
 * Build the social preview image (Open Graph / Twitter card) at
 * `apps/web/public/og.png` — 1200×630, the size every major social
 * platform crops to.
 *
 *   node scripts/build-og-image.mjs
 *
 * Reuses the playwright/chromium that's already installed for the
 * e2e suite — no new native deps. Renders an inline HTML template
 * with the brand mark, product name, tagline, key features pill row,
 * and a soft conic-gradient backdrop matching the home-page hero.
 *
 * Re-run whenever the messaging changes. The PNG is committed so we
 * don't need this in CI.
 */
import { chromium } from '@playwright/test';
import { mkdirSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = resolve(fileURLToPath(new URL('..', import.meta.url)));
const OUT_DIR = resolve(ROOT, 'apps/web/public');
mkdirSync(OUT_DIR, { recursive: true });

const html = /* html */ `
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8" />
<style>
  :root {
    --accent: #217346;
    --accent-2: #1d4ed8;
    --bg: #fafaf7;
    --fg: #0f172a;
    --muted: #475569;
  }
  * { box-sizing: border-box; }
  html, body { margin: 0; padding: 0; }
  body {
    width: 1200px;
    height: 630px;
    background: var(--bg);
    color: var(--fg);
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    position: relative;
    overflow: hidden;
    display: flex;
    flex-direction: column;
    justify-content: space-between;
    padding: 64px 72px;
  }
  .aurora {
    position: absolute; inset: -10%;
    background:
      radial-gradient(900px circle at 10% 10%, rgba(33, 115, 70, 0.28), transparent 60%),
      radial-gradient(800px circle at 95% 90%, rgba(29, 78, 216, 0.22), transparent 60%);
    filter: blur(20px);
    pointer-events: none;
  }
  .topline { display: flex; align-items: center; gap: 14px; position: relative; z-index: 1; }
  .logo {
    width: 56px; height: 56px;
    border-radius: 14px;
    background: linear-gradient(135deg, var(--accent), #2faf6b);
    display: inline-flex; align-items: center; justify-content: center;
    color: #fff;
    box-shadow: 0 8px 24px rgba(33, 115, 70, 0.25);
  }
  .logo svg { width: 32px; height: 32px; }
  .brand { font-size: 22px; font-weight: 700; letter-spacing: -0.01em; }
  .domain { font-size: 16px; color: var(--muted); margin-left: auto; font-weight: 500; }

  .main { position: relative; z-index: 1; max-width: 980px; }
  h1 {
    font-size: 80px; font-weight: 800; line-height: 1.02; margin: 0 0 18px;
    letter-spacing: -0.025em;
    background: linear-gradient(135deg, #0f172a 0%, var(--accent) 110%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
  }
  p.tagline {
    font-size: 26px; line-height: 1.36; margin: 0 0 28px; color: var(--muted);
    max-width: 920px;
    font-weight: 500;
  }
  .pills { display: flex; gap: 10px; flex-wrap: wrap; }
  .pill {
    padding: 10px 16px;
    border-radius: 999px;
    background: rgba(255,255,255,0.72);
    border: 1px solid rgba(33, 115, 70, 0.22);
    color: var(--accent);
    font-weight: 600;
    font-size: 15px;
    backdrop-filter: blur(6px);
  }

  .foot {
    position: relative; z-index: 1;
    display: flex; align-items: center; justify-content: space-between;
    color: var(--muted); font-size: 16px;
    font-weight: 500;
  }
  .foot .badges { display: flex; gap: 14px; }
  .foot .badge {
    padding: 6px 12px;
    background: rgba(15, 23, 42, 0.06);
    border-radius: 8px;
    font-family: 'JetBrains Mono', monospace;
    font-size: 14px;
    color: var(--fg);
  }
</style>
</head>
<body>
  <div class="aurora"></div>

  <div class="topline">
    <span class="logo" aria-hidden="true">
      <svg viewBox="0 0 24 24" fill="none">
        <path d="M3 3h18v18H3z" stroke="white" stroke-width="2" fill="none"/>
        <path d="M3 9h18M3 15h18M9 3v18M15 3v18" stroke="white" stroke-width="1.5"/>
      </svg>
    </span>
    <span class="brand">Casual Sheets</span>
    <span class="domain">sheet.schnsrw.live</span>
  </div>

  <div class="main">
    <h1>Excel-flavored<br/>web spreadsheet.</h1>
    <p class="tagline">Open a .xlsx in the browser, edit it like the web, save it back. Real-time co-editing, pivot tables, 8 chart types, sparklines, version history — open source and self-hostable.</p>
    <div class="pills">
      <span class="pill">.xlsx round-trip</span>
      <span class="pill">Real-time co-edit</span>
      <span class="pill">Pivot tables</span>
      <span class="pill">Charts &amp; sparklines</span>
      <span class="pill">Docker</span>
      <span class="pill">Apache-2.0</span>
    </div>
  </div>

  <div class="foot">
    <span>by Sachin Sarwa · schnsrw.live</span>
    <span class="badges">
      <span class="badge">357 e2e ✓</span>
      <span class="badge">v0.0.6</span>
    </span>
  </div>
</body>
</html>
`;

const browser = await chromium.launch();
const ctx = await browser.newContext({
  viewport: { width: 1200, height: 630 },
  deviceScaleFactor: 2, // crisp on retina + Twitter/LinkedIn upscaling
});
const page = await ctx.newPage();
await page.setContent(html, { waitUntil: 'networkidle' });
const png = await page.screenshot({ type: 'png', omitBackground: false });
await browser.close();

const outPath = resolve(OUT_DIR, 'og.png');
writeFileSync(outPath, png);
console.info(`✓ ${outPath} (${(png.byteLength / 1024).toFixed(1)} KB)`);
