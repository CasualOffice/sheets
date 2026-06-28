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

import { useCallback, useEffect, useState } from 'react';
import { adminApi, AdminApiError, getToken, setToken } from './api';
import type { AdminConfig } from './types';
import { AdminLogin } from './AdminLogin';
import { AdminLayout } from './AdminLayout';
import { BrandingSection } from './sections/BrandingSection';
import { StorageSection } from './sections/StorageSection';
import { NetworkingSection } from './sections/NetworkingSection';
import { BasePathSection } from './sections/BasePathSection';
import { LimitsSection } from './sections/LimitsSection';
import { AuthSection } from './sections/AuthSection';
import { WebhooksSection } from './sections/WebhooksSection';
import './admin.css';

/**
 * Admin panel root.
 *
 *  1. Probes /api/admin/status — when `configured: false` the panel
 *     renders an "admin not configured" hint with the env vars the
 *     operator needs to set.
 *  2. With no stored token (or invalid stored token) → AdminLogin.
 *  3. Authenticated → AdminLayout with section forms.
 *
 * Section saves are per-section (single PUT against /api/admin/config
 * with just the section payload) so a partial edit doesn't blow away
 * the other sections.
 */

type View =
  | 'loading'
  | 'not-configured'
  | 'login'
  | 'branding'
  | 'basePath'
  | 'storage'
  | 'networking'
  | 'limits'
  | 'auth'
  | 'webhooks';

export function AdminApp() {
  const [view, setView] = useState<View>('loading');
  const [config, setConfig] = useState<AdminConfig | null>(null);
  const [error, setError] = useState<string | null>(null);

  // Bootstrap: status probe → branch.
  useEffect(() => {
    (async () => {
      try {
        const status = await adminApi.status();
        if (!status.configured) {
          setView('not-configured');
          return;
        }
        if (!getToken()) {
          setView('login');
          return;
        }
        // Try to load config — if the token is stale, the server returns
        // 401 and we fall back to the login screen.
        try {
          const cfg = await adminApi.getConfig();
          setConfig(cfg);
          setView('branding');
        } catch (err) {
          if (err instanceof AdminApiError && err.status === 401) {
            setToken(null);
            setView('login');
          } else {
            setError(err instanceof Error ? err.message : 'failed to load config');
          }
        }
      } catch (err) {
        setError(err instanceof Error ? err.message : 'bootstrap failed');
      }
    })();
  }, []);

  const onLoggedIn = useCallback(async () => {
    try {
      const cfg = await adminApi.getConfig();
      setConfig(cfg);
      setView('branding');
    } catch (err) {
      setError(err instanceof Error ? err.message : 'failed to load config');
    }
  }, []);

  const onLogout = useCallback(() => {
    setToken(null);
    setConfig(null);
    setView('login');
  }, []);

  const saveSection = useCallback(
    async (patch: Partial<AdminConfig>) => {
      const next = await adminApi.putConfig(patch);
      setConfig(next);
      return next;
    },
    [],
  );

  if (view === 'loading') return <div className="admin-bootstrap">Loading admin…</div>;
  if (error) return <div className="admin-error" role="alert">{error}</div>;

  if (view === 'not-configured') {
    return (
      <div className="admin-bootstrap">
        <div className="admin-card">
          <h1>Admin panel disabled</h1>
          <p>
            Set <code>CASUAL_ADMIN_USERNAME</code>, <code>CASUAL_ADMIN_PASSWORD</code>,
            and <code>CASUAL_JWT_SECRET</code> (≥ 16 chars) on the server, then
            restart. Once those are configured this page will show a login form.
          </p>
          <p>
            See <a href="/docs/sheets/customization/" target="_blank" rel="noopener">customization docs</a> for the full guide.
          </p>
        </div>
      </div>
    );
  }

  if (view === 'login' || !config) {
    return <AdminLogin onLoggedIn={onLoggedIn} />;
  }

  const sections: { id: View; label: string; render: () => JSX.Element }[] = [
    {
      id: 'branding',
      label: 'Branding',
      render: () => <BrandingSection config={config} save={saveSection} />,
    },
    {
      id: 'basePath',
      label: 'Base path',
      render: () => <BasePathSection config={config} save={saveSection} />,
    },
    {
      id: 'storage',
      label: 'Storage',
      render: () => <StorageSection config={config} save={saveSection} />,
    },
    {
      id: 'networking',
      label: 'Networking',
      render: () => <NetworkingSection config={config} save={saveSection} />,
    },
    {
      id: 'limits',
      label: 'Room limits',
      render: () => <LimitsSection config={config} save={saveSection} />,
    },
    {
      id: 'auth',
      label: 'Auth providers',
      render: () => <AuthSection config={config} save={saveSection} />,
    },
    {
      id: 'webhooks',
      label: 'Webhooks',
      render: () => <WebhooksSection config={config} save={saveSection} />,
    },
  ];

  const active = sections.find((s) => s.id === view) ?? sections[0];

  return (
    <AdminLayout
      sections={sections.map((s) => ({ id: s.id, label: s.label }))}
      activeId={active.id}
      onNavigate={(id) => setView(id as View)}
      onLogout={onLogout}
    >
      {active.render()}
    </AdminLayout>
  );
}
