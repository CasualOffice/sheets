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

import { useEffect, useRef, useState, type KeyboardEvent } from 'react';
import { usePresence } from '../collab/presence-context';
import { Icon } from './Icon';

/**
 * Title-bar pill that surfaces the local user's display name while
 * inside a co-edit room and lets them rename inline. Mounted next to
 * the avatar stack so peers' colour bar (their own initial) and the
 * "edit my name" affordance live in the same visual cluster.
 *
 * Without this surface, the NamePrompt's "you can change this later
 * from the share menu" message was a lie — there was no "share menu"
 * for name editing. Now there's a clear discoverable affordance:
 * click your name, type, hit Enter.
 */
export function NamePill() {
  const { me, setName } = usePresence();
  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState(me?.name ?? '');
  const inputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    if (!editing) setDraft(me?.name ?? '');
  }, [me?.name, editing]);

  useEffect(() => {
    if (editing) {
      inputRef.current?.focus();
      inputRef.current?.select();
    }
  }, [editing]);

  if (!me) return null;

  const commit = () => {
    const next = draft.trim();
    if (next && next !== me.name) setName(next);
    setEditing(false);
  };

  const onKeyDown = (e: KeyboardEvent<HTMLInputElement>) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      commit();
    } else if (e.key === 'Escape') {
      e.preventDefault();
      setDraft(me.name);
      setEditing(false);
    }
  };

  if (editing) {
    return (
      <input
        ref={inputRef}
        type="text"
        className="titlebar__namepill-input"
        data-testid="titlebar-namepill-input"
        value={draft}
        maxLength={32}
        aria-label="Your display name in this room"
        onChange={(e) => setDraft(e.target.value)}
        onBlur={commit}
        onKeyDown={onKeyDown}
      />
    );
  }

  return (
    <button
      type="button"
      className="titlebar__namepill"
      data-testid="titlebar-namepill"
      title="Edit your display name"
      onClick={() => setEditing(true)}
      style={{ ['--name-color' as string]: me.color }}
    >
      <span className="titlebar__namepill-dot" aria-hidden="true" />
      <span className="titlebar__namepill-text">{me.name}</span>
      <Icon name="edit" size="sm" />
    </button>
  );
}
