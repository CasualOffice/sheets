/**
 * Lightweight, theme-aware confirm/prompt modals.
 *
 * Replaces window.confirm()/window.prompt() (native OS dialogs — jarring in the
 * Casual Office desktop shell, unstyled everywhere). Paint a small in-app modal
 * and resolve a Promise. Self-contained vanilla DOM, callable from any handler.
 */

function isDark(): boolean {
  if (typeof document !== 'undefined') {
    // Web: ThemeBridge toggles `html.univer-dark`.
    if (document.documentElement.classList.contains('univer-dark')) return true;
    const t = document.documentElement.dataset.theme;
    if (t === 'dark') return true;
    if (t === 'light') return false;
  }
  try {
    // Desktop: the launcher passes ?theme=light|dark|system.
    const tp = new URLSearchParams(window.location.search).get('theme');
    if (tp === 'dark') return true;
    if (tp === 'light') return false;
  } catch {
    /* ignore */
  }
  return typeof matchMedia === 'function' && matchMedia('(prefers-color-scheme: dark)').matches;
}

function palette() {
  return isDark()
    ? {
        bg: '#242528',
        fg: '#e9eaec',
        muted: '#a3a6ad',
        btnBg: '#33353a',
        btnBorder: '#4a4d54',
        inputBg: '#1c1d20',
        overlay: 'rgba(0,0,0,.55)',
      }
    : {
        bg: '#ffffff',
        fg: '#111111',
        muted: '#666666',
        btnBg: '#f5f5f5',
        btnBorder: '#cccccc',
        inputBg: '#ffffff',
        overlay: 'rgba(0,0,0,.45)',
      };
}

const esc = (s: string) =>
  s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]!);

function mountBackdrop(overlay: string): HTMLDivElement {
  const backdrop = document.createElement('div');
  backdrop.setAttribute(
    'style',
    `position:fixed;inset:0;z-index:2147483647;display:flex;align-items:center;justify-content:center;background:${overlay};font:14px system-ui,-apple-system,sans-serif;`,
  );
  return backdrop;
}

export interface ConfirmOptions {
  title?: string;
  body: string;
  confirmLabel?: string;
  cancelLabel?: string;
  danger?: boolean;
}

export function confirmModal(opts: ConfirmOptions): Promise<boolean> {
  return new Promise((resolve) => {
    const c = palette();
    const backdrop = mountBackdrop(c.overlay);
    const confirmBg = opts.danger ? '#dc2626' : '#2563eb';
    backdrop.innerHTML = `
      <div role="dialog" aria-modal="true" style="background:${c.bg};color:${c.fg};max-width:400px;width:90%;border-radius:12px;padding:22px 22px 16px;box-shadow:0 12px 40px rgba(0,0,0,.45);white-space:pre-line;">
        ${opts.title ? `<h2 style="margin:0 0 8px;font-size:17px;">${esc(opts.title)}</h2>` : ''}
        <p style="margin:0 0 18px;color:${c.muted};">${esc(opts.body)}</p>
        <div style="display:flex;gap:8px;justify-content:flex-end;">
          <button data-act="cancel" style="padding:8px 14px;border-radius:8px;border:1px solid ${c.btnBorder};background:${c.btnBg};color:${c.fg};cursor:pointer;">${esc(opts.cancelLabel ?? 'Cancel')}</button>
          <button data-act="ok" style="padding:8px 14px;border-radius:8px;border:0;background:${confirmBg};color:#fff;font-weight:600;cursor:pointer;">${esc(opts.confirmLabel ?? 'OK')}</button>
        </div>
      </div>`;
    document.body.appendChild(backdrop);
    const finish = (v: boolean) => {
      window.removeEventListener('keydown', onKey);
      backdrop.remove();
      resolve(v);
    };
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') finish(false);
      else if (e.key === 'Enter') finish(true);
    };
    window.addEventListener('keydown', onKey);
    backdrop.addEventListener('mousedown', (e) => {
      if (e.target === backdrop) finish(false);
    });
    backdrop.querySelector('[data-act=cancel]')!.addEventListener('click', () => finish(false));
    backdrop.querySelector('[data-act=ok]')!.addEventListener('click', () => finish(true));
    (backdrop.querySelector('[data-act=ok]') as HTMLButtonElement)?.focus();
  });
}

export interface PromptOptions {
  title?: string;
  label?: string;
  defaultValue?: string;
  confirmLabel?: string;
}

export function promptModal(opts: PromptOptions): Promise<string | null> {
  return new Promise((resolve) => {
    const c = palette();
    const backdrop = mountBackdrop(c.overlay);
    backdrop.innerHTML = `
      <div role="dialog" aria-modal="true" style="background:${c.bg};color:${c.fg};max-width:400px;width:90%;border-radius:12px;padding:22px 22px 16px;box-shadow:0 12px 40px rgba(0,0,0,.45);">
        ${opts.title ? `<h2 style="margin:0 0 8px;font-size:17px;">${esc(opts.title)}</h2>` : ''}
        ${opts.label ? `<label style="display:block;margin:0 0 6px;color:${c.muted};">${esc(opts.label)}</label>` : ''}
        <input data-act="input" type="text" style="box-sizing:border-box;width:100%;padding:9px 10px;border-radius:8px;border:1px solid ${c.btnBorder};background:${c.inputBg};color:${c.fg};margin-bottom:16px;font:inherit;" />
        <div style="display:flex;gap:8px;justify-content:flex-end;">
          <button data-act="cancel" style="padding:8px 14px;border-radius:8px;border:1px solid ${c.btnBorder};background:${c.btnBg};color:${c.fg};cursor:pointer;">Cancel</button>
          <button data-act="ok" style="padding:8px 14px;border-radius:8px;border:0;background:#2563eb;color:#fff;font-weight:600;cursor:pointer;">${esc(opts.confirmLabel ?? 'OK')}</button>
        </div>
      </div>`;
    document.body.appendChild(backdrop);
    const input = backdrop.querySelector('[data-act=input]') as HTMLInputElement;
    input.value = opts.defaultValue ?? '';
    const finish = (v: string | null) => {
      window.removeEventListener('keydown', onKey);
      backdrop.remove();
      resolve(v);
    };
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') finish(null);
      else if (e.key === 'Enter') finish(input.value);
    };
    window.addEventListener('keydown', onKey);
    backdrop.addEventListener('mousedown', (e) => {
      if (e.target === backdrop) finish(null);
    });
    backdrop.querySelector('[data-act=cancel]')!.addEventListener('click', () => finish(null));
    backdrop.querySelector('[data-act=ok]')!.addEventListener('click', () => finish(input.value));
    input.focus();
    input.select();
  });
}
