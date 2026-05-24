import { useState, type FormEvent, type ReactNode } from 'react';
import { AdminApiError } from './api';

interface Props {
  title: string;
  description: string;
  /** Submit handler; resolves on save success. Errors surface in the
   *  inline status row. */
  onSubmit: () => Promise<void>;
  children: ReactNode;
  /** Right-rail content (env-var reference / link to docs / etc). */
  aside?: ReactNode;
}

/** Common section chrome — title, description, two-column form +
 *  aside, save button + success / error status row. */
export function SectionShell({ title, description, onSubmit, children, aside }: Props) {
  const [busy, setBusy] = useState(false);
  const [status, setStatus] = useState<{ kind: 'ok' | 'err'; msg: string } | null>(null);

  const submit = async (e: FormEvent) => {
    e.preventDefault();
    setBusy(true);
    setStatus(null);
    try {
      await onSubmit();
      setStatus({ kind: 'ok', msg: 'Saved.' });
      setTimeout(() => setStatus(null), 2500);
    } catch (err) {
      const msg =
        err instanceof AdminApiError
          ? err.code ?? err.message
          : err instanceof Error
          ? err.message
          : 'Save failed.';
      setStatus({ kind: 'err', msg });
    } finally {
      setBusy(false);
    }
  };

  return (
    <form className="admin-section" onSubmit={submit}>
      <header className="admin-section__head">
        <h2 className="admin-section__title">{title}</h2>
        <p className="admin-section__desc">{description}</p>
      </header>
      <div className="admin-section__body">
        <div className="admin-section__fields">{children}</div>
        {aside && <aside className="admin-section__aside">{aside}</aside>}
      </div>
      <footer className="admin-section__foot">
        <button
          type="submit"
          className="admin-btn admin-btn--primary"
          disabled={busy}
        >
          {busy ? 'Saving…' : 'Save changes'}
        </button>
        {status && (
          <span
            className={
              status.kind === 'ok'
                ? 'admin-status admin-status--ok'
                : 'admin-status admin-status--err'
            }
            role={status.kind === 'err' ? 'alert' : undefined}
          >
            {status.msg}
          </span>
        )}
      </footer>
    </form>
  );
}
