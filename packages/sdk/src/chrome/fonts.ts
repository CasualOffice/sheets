/**
 * Ensures the chrome's webfonts (Inter for UI, Material Symbols Outlined for
 * icons) are loaded. The design system standardises on these; the rich
 * `chrome="full"` toolbar/menus render Material Symbols ligatures.
 *
 * Idempotent + safe to call repeatedly: it no-ops if the links already exist
 * (e.g. a host that already loads them, like our own app). Injected lazily only
 * when chrome is actually shown, so bare-grid SDK consumers pull nothing.
 *
 * Hosts with a strict CSP (no fonts.googleapis.com) can pre-load the fonts
 * themselves or self-host them; this is best-effort, not required for the editor
 * to function (icons fall back to the ligature text).
 */

const LINKS: Array<{ id: string; href: string }> = [
  {
    id: 'cs-font-inter',
    href: 'https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap',
  },
  {
    id: 'cs-font-material-symbols',
    href: 'https://fonts.googleapis.com/css2?family=Material+Symbols+Outlined:opsz,wght,FILL,GRAD@20..48,300..700,0..1,-50..200&display=block',
  },
];

let done = false;

export function ensureChromeFonts(): void {
  if (done || typeof document === 'undefined') return;
  done = true;
  const head = document.head;
  for (const { id, href } of LINKS) {
    if (document.getElementById(id)) continue;
    // Skip if the host already loaded the same family by any <link>.
    if (document.querySelector(`link[href*="${href.split('?')[0]}"]`)) continue;
    const link = document.createElement('link');
    link.id = id;
    link.rel = 'stylesheet';
    link.href = href;
    head.appendChild(link);
  }
}
