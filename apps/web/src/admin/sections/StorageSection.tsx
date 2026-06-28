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

const SECRET_SENTINEL = '***';

export function StorageSection({ config, save }: Props) {
  const [backend, setBackend] = useState(config.storage.backend);
  const [localPath, setLocalPath] = useState(config.storage.local.path);
  const [s3, setS3] = useState(config.storage.s3);
  const [pgUrl, setPgUrl] = useState(config.storage.postgres.url);

  const submit = async () => {
    await save({
      storage: {
        backend,
        local: { path: localPath },
        s3,
        postgres: { url: pgUrl },
      },
    });
  };

  return (
    <SectionShell
      title="Storage"
      description="Workbook persistence backend. memory = in-process (dev). local = bind-mount filesystem. s3 = AWS / MinIO / R2 / B2. postgres = bytea payload."
      onSubmit={submit}
      aside={
        <>
          <h4>Backend trade-offs</h4>
          <ul>
            <li><strong>memory</strong> — fastest, no setup; data dies on restart.</li>
            <li><strong>local</strong> — simplest persistent: bind <code>-v ./data:/data</code>; backup is rsync.</li>
            <li><strong>s3</strong> — best for horizontal scale + cross-region. Works with AWS S3, MinIO, Cloudflare R2, Backblaze B2.</li>
            <li><strong>postgres</strong> — when you already run Postgres for other apps and want one backup story.</li>
          </ul>
          <h4>Secret handling</h4>
          <p>
            The S3 secret key shows as <code>***</code> after the first save. Leave
            the sentinel in place if you don't want to change it; type the new
            value to rotate.
          </p>
        </>
      }
    >
      <label className="admin-field">
        <span>Backend</span>
        <select value={backend} onChange={(e) => setBackend(e.target.value as typeof backend)}>
          <option value="memory">memory — in-process (dev)</option>
          <option value="local">local — filesystem</option>
          <option value="s3">s3 — S3-compatible object store</option>
          <option value="postgres">postgres — bytea payload</option>
        </select>
      </label>

      {backend === 'local' && (
        <label className="admin-field">
          <span>Filesystem root</span>
          <input value={localPath} onChange={(e) => setLocalPath(e.target.value)} placeholder="/data" />
          <small>Bind-mount this path in <code>docker run -v ./workbooks:/data</code>.</small>
        </label>
      )}

      {backend === 's3' && (
        <>
          <label className="admin-field">
            <span>Endpoint <small>(blank = AWS S3)</small></span>
            <input value={s3.endpoint} onChange={(e) => setS3({ ...s3, endpoint: e.target.value })} placeholder="https://s3.amazonaws.com / http://minio:9000 / …" />
          </label>
          <label className="admin-field">
            <span>Region</span>
            <input value={s3.region} onChange={(e) => setS3({ ...s3, region: e.target.value })} placeholder="us-east-1" />
          </label>
          <label className="admin-field">
            <span>Bucket</span>
            <input value={s3.bucket} onChange={(e) => setS3({ ...s3, bucket: e.target.value })} placeholder="casual-sheets" />
          </label>
          <label className="admin-field">
            <span>Access key</span>
            <input value={s3.accessKey} onChange={(e) => setS3({ ...s3, accessKey: e.target.value })} autoComplete="off" />
          </label>
          <label className="admin-field">
            <span>Secret key</span>
            <input
              type="password"
              value={s3.secretKey}
              onChange={(e) => setS3({ ...s3, secretKey: e.target.value })}
              autoComplete="new-password"
              placeholder={s3.secretKey === SECRET_SENTINEL ? 'leave as *** to keep current' : ''}
            />
          </label>
          <label className="admin-field admin-field--check">
            <input type="checkbox" checked={s3.forcePathStyle} onChange={(e) => setS3({ ...s3, forcePathStyle: e.target.checked })} />
            <span>Force path-style addressing (MinIO + some self-hosted S3)</span>
          </label>
          <label className="admin-field">
            <span>Key prefix <small>(optional)</small></span>
            <input value={s3.keyPrefix} onChange={(e) => setS3({ ...s3, keyPrefix: e.target.value })} placeholder="prod/" />
          </label>
        </>
      )}

      {backend === 'postgres' && (
        <label className="admin-field">
          <span>Connection URL</span>
          <input type="password" value={pgUrl} onChange={(e) => setPgUrl(e.target.value)} placeholder="postgres://user:pass@host:5432/db" autoComplete="new-password" />
          <small>The <code>casual_workbooks</code> table is auto-created on first connect.</small>
        </label>
      )}
    </SectionShell>
  );
}
