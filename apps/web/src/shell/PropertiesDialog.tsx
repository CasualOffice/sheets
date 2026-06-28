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

import { useEffect, useMemo, useState } from 'react';
import { Dialog } from './Dialog';
import { useUniverAPI } from '../use-univer';
import {
  computeProperties,
  formatBytes,
  formatDate,
  readProperties,
  writeProperties,
  type WorkbookProperties,
} from './file-menu';

type Props = { onClose: () => void };

export function PropertiesDialog({ onClose }: Props) {
  const api = useUniverAPI();

  const computed = useMemo(() => (api ? computeProperties(api) : null), [api]);
  const [props, setProps] = useState<WorkbookProperties>(() => (api ? readProperties(api) : {}));

  useEffect(() => {
    if (api) setProps(readProperties(api));
  }, [api]);

  const update = (k: keyof WorkbookProperties, v: string) => setProps((p) => ({ ...p, [k]: v }));

  const save = () => {
    if (!api) return;
    writeProperties(api, {
      ...props,
      modifiedAt: new Date().toISOString(),
      createdAt: props.createdAt ?? new Date().toISOString(),
    });
    onClose();
  };

  return (
    <Dialog
      title="Properties"
      onClose={onClose}
      data-testid="properties-dialog"
      footer={
        <>
          <button type="button" className="btn-secondary" onClick={onClose}>
            Cancel
          </button>
          <button
            type="button"
            className="btn-primary"
            data-testid="properties-save"
            onClick={save}
          >
            Save
          </button>
        </>
      }
    >
      <section style={{ marginBottom: 'var(--space-6)' }}>
        <h3
          style={{
            fontSize: 'var(--text-sm)',
            color: 'var(--color-text-secondary)',
            margin: '0 0 var(--space-3)',
            fontWeight: 'var(--weight-semibold)',
            textTransform: 'uppercase',
            letterSpacing: '0.04em',
          }}
        >
          File info
        </h3>
        <ReadonlyField
          label="Name"
          testid="prop-name"
          value={computed?.name?.trim() ? computed.name : 'Untitled'}
        />
        <ReadonlyField
          label="Size"
          testid="prop-size"
          value={
            (computed && !computed.sizeIsExact ? '≈ ' : '') + formatBytes(computed?.sizeBytes ?? 0)
          }
        />
        <ReadonlyField
          label="Sheets"
          testid="prop-sheets"
          value={String(computed?.sheetCount ?? 0)}
        />
        <ReadonlyField
          label="Cells with data"
          testid="prop-cells"
          value={String(computed?.cellCount ?? 0)}
        />
        <ReadonlyField label="Created" testid="prop-created" value={formatDate(props.createdAt)} />
        <ReadonlyField
          label="Last modified"
          testid="prop-modified"
          value={formatDate(props.modifiedAt)}
        />
      </section>

      <section>
        <h3
          style={{
            fontSize: 'var(--text-sm)',
            color: 'var(--color-text-secondary)',
            margin: '0 0 var(--space-3)',
            fontWeight: 'var(--weight-semibold)',
            textTransform: 'uppercase',
            letterSpacing: '0.04em',
          }}
        >
          Metadata
        </h3>
        <EditableField
          label="Title"
          testid="prop-title"
          value={props.title ?? ''}
          onChange={(v) => update('title', v)}
        />
        <EditableField
          label="Subject"
          testid="prop-subject"
          value={props.subject ?? ''}
          onChange={(v) => update('subject', v)}
        />
        <EditableField
          label="Author"
          testid="prop-author"
          value={props.author ?? ''}
          onChange={(v) => update('author', v)}
        />
        <EditableField
          label="Tags"
          testid="prop-tags"
          value={props.tags ?? ''}
          onChange={(v) => update('tags', v)}
        />
        <EditableField
          label="Category"
          testid="prop-category"
          value={props.category ?? ''}
          onChange={(v) => update('category', v)}
        />
        <EditableField
          label="Description"
          testid="prop-description"
          textarea
          value={props.description ?? ''}
          onChange={(v) => update('description', v)}
        />
      </section>
    </Dialog>
  );
}

function ReadonlyField({ label, value, testid }: { label: string; value: string; testid: string }) {
  return (
    <div className="field">
      <span className="field__label">{label}</span>
      <span className="field__value field__value--mono" data-testid={testid}>
        {value}
      </span>
    </div>
  );
}

function EditableField({
  label,
  value,
  onChange,
  textarea,
  testid,
}: {
  label: string;
  value: string;
  onChange: (v: string) => void;
  textarea?: boolean;
  testid: string;
}) {
  return (
    <div className="field">
      <label className="field__label" htmlFor={testid}>
        {label}
      </label>
      {textarea ? (
        <textarea
          id={testid}
          className="input input--textarea"
          data-testid={testid}
          value={value}
          rows={3}
          onChange={(e) => onChange(e.target.value)}
        />
      ) : (
        <input
          id={testid}
          className="input"
          data-testid={testid}
          value={value}
          onChange={(e) => onChange(e.target.value)}
        />
      )}
    </div>
  );
}
