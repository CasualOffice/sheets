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

import { useState } from 'react';
import type { AdminConfig } from '../types';
import { SectionShell } from '../SectionShell';

interface Props {
  config: AdminConfig;
  save: (patch: Partial<AdminConfig>) => Promise<AdminConfig>;
}

export function BrandingSection({ config, save }: Props) {
  const [appName, setAppName] = useState(config.branding.appName);
  const [accentColor, setAccentColor] = useState(config.branding.accentColor);
  const [logoUrl, setLogoUrl] = useState(config.branding.logoUrl ?? '');

  const submit = async () => {
    await save({
      branding: {
        appName,
        accentColor,
        logoUrl: logoUrl.trim() === '' ? null : logoUrl,
      },
    });
  };

  return (
    <SectionShell
      title="Branding"
      description="App name + accent colour + logo. Drives the title bar, OG image alt, and the --color-accent CSS variable across the editor."
      onSubmit={submit}
      aside={
        <>
          <h4>Tips</h4>
          <ul>
            <li>The accent colour also tints the toolbar selection highlight, the focus ring, and the formula bar's editing underline.</li>
            <li>Logo can be a relative URL (e.g. <code>/logo.svg</code>) served from your reverse proxy or a full <code>https://</code> URL.</li>
            <li>Per-deployment branding ships in v0.1; <em>per-user</em> branding (theme override from the JWT) lands in v0.2.</li>
          </ul>
        </>
      }
    >
      <label className="admin-field">
        <span>App name</span>
        <input value={appName} onChange={(e) => setAppName(e.target.value)} maxLength={48} />
      </label>
      <label className="admin-field">
        <span>Accent colour</span>
        <div className="admin-field__row">
          <input
            type="color"
            value={accentColor}
            onChange={(e) => setAccentColor(e.target.value)}
            className="admin-field__color"
          />
          <input
            type="text"
            value={accentColor}
            onChange={(e) => setAccentColor(e.target.value)}
            pattern="^#[0-9a-fA-F]{6}$"
            maxLength={7}
          />
        </div>
      </label>
      <label className="admin-field">
        <span>Logo URL</span>
        <input
          type="text"
          value={logoUrl}
          onChange={(e) => setLogoUrl(e.target.value)}
          placeholder="/logo.svg or https://…"
        />
      </label>
    </SectionShell>
  );
}
