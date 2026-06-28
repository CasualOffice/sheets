/**
 * Copyright 2026 Casual Office
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

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
