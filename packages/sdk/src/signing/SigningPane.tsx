/**
 * SigningPane — the floating sidebar that walks the signer through
 * fields. Lives inside <SigningProvider>; uses the controller via
 * `useSigning()`.
 *
 * Layout: a right-anchored panel showing
 *   - Banner (optional, from session config)
 *   - Field list with state markers (active / signed / pending)
 *   - Method picker for the active field
 *   - Capture surface (Drawn / Typed / Uploaded) depending on method
 *   - Footer: Cancel + Complete
 *
 * The pane is presentation-only — every action routes through the
 * SigningProvider's helpers (signField, completeIfReady, cancel)
 * so the controller stays the single source of truth.
 */

import { useEffect, useState, type CSSProperties } from 'react';

import { useSigning } from './SigningProvider';
import type { SignatureField, SignatureMethod, SignedFieldPayload } from './types';
import {
  DrawnSignaturePad,
  TypedSignatureField,
  UploadedSignatureField,
  type CapturedSignature,
} from './captures';

export interface SigningPaneProps {
  /** Optional banner override; falls back to session.banner. */
  banner?: string;
  /** Optional data-testid root. */
  testId?: string;
}

export function SigningPane({ banner, testId = 'signing-pane' }: SigningPaneProps) {
  const ctx = useSigning();
  if (!ctx) return null;

  const { snapshot, signField, completeIfReady, cancel } = ctx;
  if (snapshot.isComplete || snapshot.isCancelled) return null;

  const active: SignatureField | null =
    snapshot.activeFieldIndex >= 0 ? snapshot.fields[snapshot.activeFieldIndex] : null;

  return (
    <aside style={paneStyle} role="region" aria-label="Signing pane" data-testid={testId}>
      {banner && (
        <div style={bannerStyle} data-testid={`${testId}-banner`}>
          {banner}
        </div>
      )}
      <div style={listStyle} data-testid={`${testId}-fields`}>
        {snapshot.fields.map((f, i) => {
          const isSigned = !!snapshot.signed[f.fieldId];
          const isActive = i === snapshot.activeFieldIndex;
          return (
            <div
              key={f.fieldId}
              style={listItemStyle(isActive, isSigned)}
              data-testid={`${testId}-field-${f.fieldId}`}
              data-state={isSigned ? 'signed' : isActive ? 'active' : 'pending'}
            >
              <span style={listIconStyle(isSigned)} aria-hidden="true">
                {isSigned ? '✓' : i + 1}
              </span>
              <span style={listLabelStyle}>{f.label}</span>
              {!f.required && (
                <span style={optionalChipStyle} aria-label="Optional">
                  optional
                </span>
              )}
            </div>
          );
        })}
      </div>

      {active && (
        <ActiveFieldEditor
          field={active}
          testId={testId}
          onCapture={async (cap, method) => {
            const payload: SignedFieldPayload = {
              fieldId: active.fieldId,
              method,
              bytes: cap.bytes,
              mime: cap.mime,
              signedAt: new Date().toISOString(),
            };
            await signField(payload);
          }}
        />
      )}

      {!active && snapshot.canComplete && (
        <div style={completeBlockStyle} data-testid={`${testId}-complete-block`}>
          All required signatures collected. Ready to finalise.
        </div>
      )}

      <footer style={footerStyle}>
        <button
          type="button"
          onClick={() => cancel('signer_cancelled')}
          style={secondaryBtnStyle()}
          data-testid={`${testId}-cancel`}
        >
          Cancel
        </button>
        <button
          type="button"
          onClick={() => void completeIfReady()}
          disabled={!snapshot.canComplete}
          style={primaryBtnStyle(!snapshot.canComplete)}
          data-testid={`${testId}-complete`}
        >
          Complete
        </button>
      </footer>
    </aside>
  );
}

// ---------------------------------------------------------------
// Active field editor — method picker + capture surface
// ---------------------------------------------------------------

function ActiveFieldEditor({
  field,
  testId,
  onCapture,
}: {
  field: SignatureField;
  testId: string;
  onCapture: (cap: CapturedSignature, method: SignatureMethod) => void | Promise<void>;
}) {
  const [method, setMethod] = useState<SignatureMethod>(field.methods[0]);

  // Reset the method picker whenever the active field changes — a
  // method that was valid for the previous field may not be in this
  // field's list.
  useEffect(() => {
    setMethod(field.methods[0]);
  }, [field]);

  return (
    <div style={editorStyle} data-testid={`${testId}-editor`}>
      <div style={editorHeaderStyle}>
        <div style={editorLabelStyle}>{field.label}</div>
        {field.signer?.name && <div style={editorSignerStyle}>{field.signer.name}</div>}
      </div>
      {field.methods.length > 1 && (
        <div style={methodTabsStyle} role="tablist" data-testid={`${testId}-methods`}>
          {field.methods.map((m) => (
            <button
              key={m}
              type="button"
              role="tab"
              aria-selected={method === m}
              onClick={() => setMethod(m)}
              style={methodTabStyle(method === m)}
              data-testid={`${testId}-method-${m}`}
            >
              {methodLabel(m)}
            </button>
          ))}
        </div>
      )}
      <div style={captureWrapStyle}>
        {method === 'drawn' && <DrawnSignaturePad onCapture={(c) => onCapture(c, 'drawn')} />}
        {method === 'typed' && (
          <TypedSignatureField
            defaultText={field.signer?.name ?? ''}
            onCapture={(c) => onCapture(c, 'typed')}
          />
        )}
        {method === 'uploaded' && (
          <UploadedSignatureField onCapture={(c) => onCapture(c, 'uploaded')} />
        )}
      </div>
    </div>
  );
}

function methodLabel(m: SignatureMethod): string {
  switch (m) {
    case 'drawn':
      return 'Draw';
    case 'typed':
      return 'Type';
    case 'uploaded':
      return 'Upload';
  }
}

// ---------------------------------------------------------------
// Styles
// ---------------------------------------------------------------

const paneStyle: CSSProperties = {
  position: 'fixed',
  top: 16,
  right: 16,
  bottom: 16,
  width: 360,
  maxWidth: '100vw',
  display: 'flex',
  flexDirection: 'column',
  gap: 14,
  padding: 16,
  background: 'var(--doc-surface, #fff)',
  border: '1px solid var(--doc-border, #cbd5e1)',
  borderRadius: 12,
  boxShadow: '0 1px 1px rgba(0, 0, 0, 0.04), 0 6px 24px rgba(15, 23, 42, 0.12)',
  fontFamily: 'inherit',
  zIndex: 9000,
};

const bannerStyle: CSSProperties = {
  padding: '8px 10px',
  background: 'var(--doc-surface-2, #f1f5f9)',
  border: '1px solid var(--doc-border-light, #e2e8f0)',
  borderRadius: 6,
  fontSize: 12,
  color: 'var(--doc-text-muted, #475569)',
};

const listStyle: CSSProperties = {
  display: 'flex',
  flexDirection: 'column',
  gap: 4,
};

function listItemStyle(active: boolean, signed: boolean): CSSProperties {
  return {
    display: 'flex',
    alignItems: 'center',
    gap: 10,
    padding: '8px 10px',
    borderRadius: 6,
    background: active ? 'var(--doc-surface-2, #f1f5f9)' : signed ? 'transparent' : 'transparent',
    border: active ? '1px solid var(--doc-border, #cbd5e1)' : '1px solid transparent',
    opacity: signed && !active ? 0.7 : 1,
  };
}

function listIconStyle(signed: boolean): CSSProperties {
  return {
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
    width: 22,
    height: 22,
    borderRadius: '50%',
    background: signed ? 'var(--doc-accent, #2563eb)' : 'var(--doc-surface-2, #f1f5f9)',
    color: signed ? '#fff' : 'var(--doc-text-muted, #475569)',
    fontSize: 12,
    fontWeight: 600,
    flexShrink: 0,
  };
}

const listLabelStyle: CSSProperties = {
  flex: 1,
  fontSize: 13,
  color: 'var(--doc-text, #0f172a)',
  fontWeight: 500,
};

const optionalChipStyle: CSSProperties = {
  fontSize: 11,
  color: 'var(--doc-text-muted, #64748b)',
  padding: '2px 6px',
  background: 'var(--doc-surface-2, #f1f5f9)',
  borderRadius: 4,
};

const editorStyle: CSSProperties = {
  display: 'flex',
  flexDirection: 'column',
  gap: 12,
};

const editorHeaderStyle: CSSProperties = {
  display: 'flex',
  flexDirection: 'column',
  gap: 2,
};

const editorLabelStyle: CSSProperties = {
  fontSize: 13,
  fontWeight: 600,
  color: 'var(--doc-text, #0f172a)',
};

const editorSignerStyle: CSSProperties = {
  fontSize: 12,
  color: 'var(--doc-text-muted, #64748b)',
};

const methodTabsStyle: CSSProperties = {
  display: 'flex',
  gap: 4,
  padding: 2,
  background: 'var(--doc-surface-2, #f1f5f9)',
  borderRadius: 6,
};

function methodTabStyle(selected: boolean): CSSProperties {
  return {
    flex: 1,
    padding: '6px 10px',
    background: selected ? 'var(--doc-surface, #fff)' : 'transparent',
    border: 'none',
    borderRadius: 4,
    fontSize: 12,
    fontWeight: 500,
    color: selected ? 'var(--doc-text, #0f172a)' : 'var(--doc-text-muted, #475569)',
    cursor: 'pointer',
    fontFamily: 'inherit',
  };
}

const captureWrapStyle: CSSProperties = {
  display: 'flex',
  flexDirection: 'column',
  gap: 8,
};

const completeBlockStyle: CSSProperties = {
  padding: '10px 12px',
  background: 'rgba(34, 197, 94, 0.08)',
  border: '1px solid rgba(34, 197, 94, 0.28)',
  borderRadius: 6,
  fontSize: 12,
  color: 'rgb(20, 83, 45)',
};

const footerStyle: CSSProperties = {
  display: 'flex',
  justifyContent: 'flex-end',
  gap: 8,
  marginTop: 'auto',
  paddingTop: 12,
  borderTop: '1px solid var(--doc-border-light, #e2e8f0)',
};

function primaryBtnStyle(disabled: boolean): CSSProperties {
  return {
    padding: '8px 16px',
    borderRadius: 6,
    border: '1px solid transparent',
    background: disabled ? 'var(--doc-border, #cbd5e1)' : 'var(--doc-accent, #2563eb)',
    color: disabled ? 'var(--doc-text-muted, #64748b)' : '#fff',
    fontSize: 13,
    fontWeight: 600,
    cursor: disabled ? 'not-allowed' : 'pointer',
    fontFamily: 'inherit',
  };
}

function secondaryBtnStyle(): CSSProperties {
  return {
    padding: '8px 16px',
    borderRadius: 6,
    border: '1px solid var(--doc-border, #cbd5e1)',
    background: 'transparent',
    color: 'var(--doc-text, #0f172a)',
    fontSize: 13,
    fontWeight: 500,
    cursor: 'pointer',
    fontFamily: 'inherit',
  };
}
