import { useState } from 'react';
import type { AdminConfig } from '../types';
import { SectionShell } from '../SectionShell';

interface Props {
  config: AdminConfig;
  save: (patch: Partial<AdminConfig>) => Promise<AdminConfig>;
}

export function BasePathSection({ config, save }: Props) {
  const [basePath, setBasePath] = useState(config.basePath);

  const submit = async () => {
    // Normalise: strip trailing slash, ensure leading slash.
    let v = basePath.trim();
    if (v === '' || v === '/') v = '';
    else {
      if (!v.startsWith('/')) v = '/' + v;
      if (v.endsWith('/')) v = v.slice(0, -1);
    }
    setBasePath(v);
    await save({ basePath: v });
  };

  return (
    <SectionShell
      title="Base path"
      description="Reverse-proxy mount path. Empty = served at /. Setting this prefixes every server route + adjusts the SPA's asset base."
      onSubmit={submit}
      aside={
        <>
          <h4>When to set this</h4>
          <p>
            You're running Casual Sheets behind a reverse proxy at a path like{' '}
            <code>https://acme.example/sheets</code> (sharing the host with other
            apps). The base path tells the server to expect inbound URLs prefixed
            with <code>/sheets</code> and emit asset references the same way.
          </p>
          <h4>What to update upstream</h4>
          <ul>
            <li>Reverse proxy: forward <code>/sheets/*</code> + <code>/sheets/yjs</code> (WebSocket upgrade) to this container.</li>
            <li>Don't strip the prefix on the way through — the server expects to see it.</li>
            <li>Bump <code>CASUAL_PUBLIC_ORIGIN</code> to the externally-visible URL including the prefix.</li>
          </ul>
        </>
      }
    >
      <label className="admin-field">
        <span>Mount path</span>
        <input
          value={basePath}
          onChange={(e) => setBasePath(e.target.value)}
          placeholder="/sheets"
        />
        <small>
          Leading slash, no trailing slash. Leave empty to serve at <code>/</code>.
        </small>
      </label>
    </SectionShell>
  );
}
