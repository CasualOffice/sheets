import { useState } from 'react';
import type { AdminConfig } from '../types';
import { SectionShell } from '../SectionShell';

interface Props {
  config: AdminConfig;
  save: (patch: Partial<AdminConfig>) => Promise<AdminConfig>;
}

export function LimitsSection({ config, save }: Props) {
  const [maxRooms, setMaxRooms] = useState(config.limits.maxRooms);
  const [maxFile, setMaxFile] = useState(config.limits.maxFileSizeMb);
  const [ttl, setTtl] = useState(config.limits.roomTtlMin);
  const [maxUsers, setMaxUsers] = useState(config.limits.maxUsersPerRoom);

  const submit = async () => {
    await save({
      limits: {
        maxRooms,
        maxFileSizeMb: maxFile,
        roomTtlMin: ttl,
        maxUsersPerRoom: maxUsers,
      },
    });
  };

  return (
    <SectionShell
      title="Room limits"
      description="Operator-visible policy knobs. Useful for shared deployments — bounds memory + bandwidth."
      onSubmit={submit}
      aside={
        <>
          <h4>Sizing rule of thumb</h4>
          <ul>
            <li>One idle room ≈ 50–500 KB of resident memory.</li>
            <li>Active editing (100 mutations/min) adds ~1–2 MB transient.</li>
            <li>For 100 concurrent rooms with light editing, 256 MB of RSS is enough headroom.</li>
          </ul>
        </>
      }
    >
      <label className="admin-field">
        <span>Max rooms <small>(simultaneous)</small></span>
        <input type="number" min={1} value={maxRooms} onChange={(e) => setMaxRooms(Number(e.target.value) || 1)} />
      </label>
      <label className="admin-field">
        <span>Max file size <small>(MiB)</small></span>
        <input type="number" min={1} value={maxFile} onChange={(e) => setMaxFile(Number(e.target.value) || 1)} />
      </label>
      <label className="admin-field">
        <span>Room TTL <small>(minutes after last client disconnects)</small></span>
        <input type="number" min={1} value={ttl} onChange={(e) => setTtl(Number(e.target.value) || 1)} />
      </label>
      <label className="admin-field">
        <span>Max users per room</span>
        <input type="number" min={1} value={maxUsers} onChange={(e) => setMaxUsers(Number(e.target.value) || 1)} />
      </label>
    </SectionShell>
  );
}
