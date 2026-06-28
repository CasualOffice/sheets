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

export function NetworkingSection({ config, save }: Props) {
  const [publicOrigin, setPublicOrigin] = useState(config.networking.publicOrigin);
  const [corsOrigins, setCorsOrigins] = useState(config.networking.corsOrigins);
  const [trustProxy, setTrustProxy] = useState(config.networking.trustProxy);
  const [hsts, setHsts] = useState(config.networking.hstsMaxAge);

  const submit = async () => {
    await save({
      networking: { publicOrigin, corsOrigins, trustProxy, hstsMaxAge: hsts },
    });
  };

  return (
    <SectionShell
      title="Networking"
      description="Reverse-proxy + CORS + HSTS settings for production deployments."
      onSubmit={submit}
      aside={
        <>
          <h4>Public origin</h4>
          <p>
            The externally-visible URL of this deployment. Used by WOPI
            <code>BaseFileName</code>, share-link generation, and the
            <code>og:url</code> meta tag.
          </p>
          <h4>CORS</h4>
          <p>
            Empty list = same-origin only (recommended). Set when your admin
            panel or API is called from a different origin than the SPA.
          </p>
          <h4>HSTS</h4>
          <p>
            Only emit when HTTPS terminates upstream. Sending the header over
            HTTP can lock users out for the max-age window.
          </p>
        </>
      }
    >
      <label className="admin-field">
        <span>Public origin</span>
        <input value={publicOrigin} onChange={(e) => setPublicOrigin(e.target.value)} placeholder="https://sheets.acme.example" />
      </label>
      <label className="admin-field">
        <span>CORS origins <small>(comma-separated)</small></span>
        <input value={corsOrigins} onChange={(e) => setCorsOrigins(e.target.value)} placeholder="https://app.acme.example,https://staging.acme.example" />
      </label>
      <label className="admin-field">
        <span>Trust proxy</span>
        <input value={trustProxy} onChange={(e) => setTrustProxy(e.target.value)} placeholder="loopback" />
        <small>
          <code>loopback</code> trusts 127.0.0.1; <code>true</code> trusts the immediate hop; a list of IPs / CIDRs is explicit.
        </small>
      </label>
      <label className="admin-field">
        <span>HSTS max-age <small>(seconds; 0 = disabled)</small></span>
        <input type="number" min={0} value={hsts} onChange={(e) => setHsts(Number(e.target.value) || 0)} />
      </label>
    </SectionShell>
  );
}
