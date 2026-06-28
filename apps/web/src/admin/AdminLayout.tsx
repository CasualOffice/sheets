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

import type { ReactNode } from 'react';

interface SectionMeta {
  id: string;
  label: string;
}

interface Props {
  sections: SectionMeta[];
  activeId: string;
  onNavigate: (id: string) => void;
  onLogout: () => void;
  children: ReactNode;
}

export function AdminLayout({
  sections,
  activeId,
  onNavigate,
  onLogout,
  children,
}: Props) {
  return (
    <div className="admin">
      <aside className="admin__side">
        <div className="admin__brand">
          <strong>Casual Sheets</strong>
          <span>Admin</span>
        </div>
        <nav className="admin__nav" aria-label="Admin sections">
          {sections.map((s) => (
            <button
              key={s.id}
              type="button"
              className={
                s.id === activeId
                  ? 'admin__nav-item admin__nav-item--active'
                  : 'admin__nav-item'
              }
              onClick={() => onNavigate(s.id)}
            >
              {s.label}
            </button>
          ))}
        </nav>
        <div className="admin__side-foot">
          <button
            type="button"
            className="admin-btn admin-btn--ghost"
            onClick={onLogout}
          >
            Sign out
          </button>
          <a
            className="admin__doc-link"
            href="/docs/sheets/customization/"
            target="_blank"
            rel="noopener"
          >
            Customization docs →
          </a>
        </div>
      </aside>
      <main className="admin__main">{children}</main>
    </div>
  );
}
