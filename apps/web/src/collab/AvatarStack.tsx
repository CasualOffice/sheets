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

import { Tooltip } from '../shell/Tooltip';
import { initials } from './presence';
import { usePresence } from './presence-context';
import { useCollab } from './collab-context';

/**
 * Compact avatar stack rendered in the titlebar when joined to a room.
 * Each peer is a colored circle with their initials; tooltip shows the
 * full name + last-seen heartbeat. Limits to 4 visible avatars + "+N"
 * pill so a packed room doesn't blow out the header. Self goes first.
 */
const VISIBLE = 4;
const IDLE_AFTER_MS = 8_000;

export function AvatarStack() {
  const { roomId, status } = useCollab();
  const { peers, me } = usePresence();
  if (!roomId || status === 'off') return null;

  type Entry = { clientId: number; name: string; color: string; lastSeen?: number };
  const all: Entry[] = me
    ? [
        { clientId: -1, name: me.name, color: me.color, lastSeen: Date.now() },
        ...peers.map((p) => ({
          clientId: p.clientId,
          name: p.name,
          color: p.color,
          lastSeen: p.lastSeen,
        })),
      ]
    : peers.map((p) => ({
        clientId: p.clientId,
        name: p.name,
        color: p.color,
        lastSeen: p.lastSeen,
      }));

  if (all.length === 0) return null;

  const visible = all.slice(0, VISIBLE);
  const extra = all.length - visible.length;
  const now = Date.now();

  return (
    <span className="presence-avatars" data-testid="presence-avatars">
      {visible.map((p) => {
        const isSelf = p.clientId === -1;
        const ago = p.lastSeen ? now - p.lastSeen : 0;
        const idle = !isSelf && ago > IDLE_AFTER_MS;
        const label = isSelf
          ? `${p.name} (you)`
          : `${p.name} · ${formatLastSeen(ago)}`;
        return (
          <Tooltip key={p.clientId} label={label} side="bottom">
            <span
              className={'presence-avatar' + (idle ? ' presence-avatar--idle' : '')}
              data-testid="presence-avatar"
              data-idle={idle ? '1' : '0'}
              style={{ background: p.color }}
              aria-label={label}
            >
              {initials(p.name)}
            </span>
          </Tooltip>
        );
      })}
      {extra > 0 && (
        <Tooltip label={`${extra} more in this room`} side="bottom">
          <span className="presence-avatar presence-avatar--more" data-testid="presence-avatar-more">
            +{extra}
          </span>
        </Tooltip>
      )}
    </span>
  );
}

function formatLastSeen(ms: number): string {
  if (ms < IDLE_AFTER_MS) return 'Active now';
  const sec = Math.floor(ms / 1000);
  if (sec < 60) return `Last seen ${sec}s ago`;
  const min = Math.floor(sec / 60);
  if (min < 60) return `Last seen ${min}m ago`;
  const hr = Math.floor(min / 60);
  return `Last seen ${hr}h ago`;
}
