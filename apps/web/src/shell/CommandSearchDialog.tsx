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

import { useEffect, useMemo, useRef, useState } from 'react';
import { Dialog } from './Dialog';

export type CommandSearchItem = {
  id: string;
  label: string;
  path: string;
  shortcut?: string;
  run: () => void | Promise<void>;
};

type Props = {
  items: CommandSearchItem[];
  onClose: () => void;
};

export function CommandSearchDialog({ items, onClose }: Props) {
  const [query, setQuery] = useState('');
  const [activeIndex, setActiveIndex] = useState(0);
  const inputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    inputRef.current?.focus();
    inputRef.current?.select();
  }, []);

  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) return items;
    return items.filter((item) => {
      const haystack = [item.label, item.path, item.shortcut ?? '', item.id].join(' ').toLowerCase();
      return haystack.includes(q);
    });
  }, [items, query]);

  const currentIndex = filtered.length === 0 ? -1 : Math.min(activeIndex, filtered.length - 1);
  const current = currentIndex >= 0 ? filtered[currentIndex] : null;

  useEffect(() => {
    setActiveIndex(0);
  }, [query]);

  const run = async (item: CommandSearchItem | null) => {
    if (!item) return;
    await item.run();
    onClose();
  };

  return (
    <Dialog title="Tell Me" onClose={onClose} data-testid="command-search-dialog">
      <div className="command-search">
        <div className="field">
          <label className="field__label" htmlFor="command-search-input">
            Search commands
          </label>
          <input
            ref={inputRef}
            id="command-search-input"
            className="input"
            data-testid="command-search-input"
            placeholder="Type to search menu commands"
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'ArrowDown') {
                e.preventDefault();
                if (filtered.length > 0) setActiveIndex((i) => Math.min(i + 1, filtered.length - 1));
              } else if (e.key === 'ArrowUp') {
                e.preventDefault();
                if (filtered.length > 0) setActiveIndex((i) => Math.max(i - 1, 0));
              } else if (e.key === 'Home') {
                e.preventDefault();
                setActiveIndex(0);
              } else if (e.key === 'End') {
                e.preventDefault();
                setActiveIndex(Math.max(0, filtered.length - 1));
              } else if (e.key === 'Enter') {
                e.preventDefault();
                void run(current);
              }
            }}
          />
        </div>

        <div className="command-search__results" role="listbox" aria-label="Matching commands">
          {filtered.length === 0 ? (
            <div className="command-search__empty" data-testid="command-search-empty">
              No matching commands.
            </div>
          ) : (
            filtered.map((item, index) => (
              <button
                key={item.id}
                type="button"
                className={`command-search__item${index === currentIndex ? ' command-search__item--active' : ''}`}
                data-testid={`command-search-item-${item.id}`}
                onMouseEnter={() => setActiveIndex(index)}
                onClick={() => void run(item)}
              >
                <span className="command-search__main">
                  <span className="command-search__label">{item.label}</span>
                  <span className="command-search__path">{item.path}</span>
                </span>
                {item.shortcut ? <span className="command-search__shortcut">{item.shortcut}</span> : null}
              </button>
            ))
          )}
        </div>

        <p className="command-search__hint">
          Use Arrow keys to move, Enter to run, Esc to close.
        </p>
      </div>
    </Dialog>
  );
}
