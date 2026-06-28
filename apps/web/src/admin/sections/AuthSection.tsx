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

export function AuthSection({ config, save }: Props) {
  const [oidc, setOidc] = useState(config.auth.oidc);
  const [saml, setSaml] = useState(config.auth.saml);
  const [jwt, setJwt] = useState(config.auth.jwt);

  const submit = async () => {
    await save({ auth: { oidc, saml, jwt } });
  };

  return (
    <SectionShell
      title="Auth providers"
      description="JWT is live in v0.1; OIDC + SAML are configurable here but enforced in v0.2."
      onSubmit={submit}
      aside={
        <>
          <h4>JWT (live)</h4>
          <p>
            Set <code>CASUAL_JWT_SECRET</code> on the server to enable. Tokens
            are minted via <code>POST /api/tokens</code> (admin role required).
            See <a href="/docs/sheets/customization/" target="_blank" rel="noopener">customization docs</a>.
          </p>
          <h4>OIDC + SAML (stub)</h4>
          <p>
            The forms here persist into the config so v0.2 can ship the
            enforcement code without breaking on-disk configs from v0.1.
          </p>
        </>
      }
    >
      <fieldset className="admin-fieldset">
        <legend>JWT</legend>
        <label className="admin-field admin-field--check">
          <input type="checkbox" checked={jwt.enabled} onChange={(e) => setJwt({ ...jwt, enabled: e.target.checked })} />
          <span>Enable JWT auth on WOPI routes (also requires <code>CASUAL_JWT_SECRET</code> on the server)</span>
        </label>
        <label className="admin-field">
          <span>Issuer <small>(token <code>iss</code> claim)</small></span>
          <input value={jwt.issuer} onChange={(e) => setJwt({ ...jwt, issuer: e.target.value })} placeholder="https://sheets.acme.example" />
        </label>
        <label className="admin-field">
          <span>Default TTL <small>(seconds)</small></span>
          <input type="number" min={60} value={jwt.defaultTtlSeconds} onChange={(e) => setJwt({ ...jwt, defaultTtlSeconds: Number(e.target.value) || 3600 })} />
        </label>
      </fieldset>

      <fieldset className="admin-fieldset">
        <legend>OIDC <small>(stub — backend in v0.2)</small></legend>
        <label className="admin-field admin-field--check">
          <input type="checkbox" checked={oidc.enabled} onChange={(e) => setOidc({ ...oidc, enabled: e.target.checked })} />
          <span>Enable OIDC sign-in</span>
        </label>
        <label className="admin-field">
          <span>Issuer URL</span>
          <input value={oidc.issuer} onChange={(e) => setOidc({ ...oidc, issuer: e.target.value })} placeholder="https://login.acme.example" />
        </label>
        <label className="admin-field">
          <span>Client ID</span>
          <input value={oidc.clientId} onChange={(e) => setOidc({ ...oidc, clientId: e.target.value })} />
        </label>
        <label className="admin-field">
          <span>Client secret</span>
          <input type="password" value={oidc.clientSecret} onChange={(e) => setOidc({ ...oidc, clientSecret: e.target.value })} autoComplete="new-password" placeholder={oidc.clientSecret === '***' ? 'leave as *** to keep current' : ''} />
        </label>
        <label className="admin-field">
          <span>Redirect URI</span>
          <input value={oidc.redirectUri} onChange={(e) => setOidc({ ...oidc, redirectUri: e.target.value })} placeholder="https://sheets.acme.example/auth/oidc/callback" />
        </label>
      </fieldset>

      <fieldset className="admin-fieldset">
        <legend>SAML <small>(stub — backend in v0.2)</small></legend>
        <label className="admin-field admin-field--check">
          <input type="checkbox" checked={saml.enabled} onChange={(e) => setSaml({ ...saml, enabled: e.target.checked })} />
          <span>Enable SAML sign-in</span>
        </label>
        <label className="admin-field">
          <span>IdP metadata URL</span>
          <input value={saml.idpMetadataUrl} onChange={(e) => setSaml({ ...saml, idpMetadataUrl: e.target.value })} />
        </label>
        <label className="admin-field">
          <span>SP entity ID</span>
          <input value={saml.spEntityId} onChange={(e) => setSaml({ ...saml, spEntityId: e.target.value })} />
        </label>
      </fieldset>
    </SectionShell>
  );
}
