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

/**
 * Three signature-capture surfaces: drawn (canvas), typed (input
 * rendered in a script font), uploaded (image file picker).
 *
 * Each one emits an `{ bytes, mime }` pair that the parent feeds
 * into SigningContext.signField. The components are presentation +
 * input only — they don't know about the controller or the field
 * they're capturing for.
 */

import { useEffect, useRef, useState, type CSSProperties } from 'react';

export interface CapturedSignature {
  bytes: ArrayBuffer;
  mime: string;
}

// ---------------------------------------------------------------
// Drawn — canvas
// ---------------------------------------------------------------

export interface DrawnSignaturePadProps {
  /** Fired when the user clicks "Use this signature". */
  onCapture: (sig: CapturedSignature) => void;
  /** Optional clear-button label override. */
  clearLabel?: string;
  /** Optional save-button label override. */
  saveLabel?: string;
  /** Canvas pixel size. Default 480 × 160. */
  width?: number;
  height?: number;
}

export function DrawnSignaturePad({
  onCapture,
  clearLabel = 'Clear',
  saveLabel = 'Use this signature',
  width = 480,
  height = 160,
}: DrawnSignaturePadProps) {
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const drawingRef = useRef(false);
  const [hasInk, setHasInk] = useState(false);

  // Set up a clean canvas on mount. Background is transparent so
  // the stamped image composites cleanly over the document.
  useEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    ctx.strokeStyle = '#0f172a';
    ctx.lineWidth = 2;
    ctx.lineCap = 'round';
    ctx.lineJoin = 'round';
  }, []);

  const start = (e: React.PointerEvent<HTMLCanvasElement>) => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    const { x, y } = pointerPos(e, canvas);
    ctx.beginPath();
    ctx.moveTo(x, y);
    drawingRef.current = true;
    canvas.setPointerCapture(e.pointerId);
  };

  const move = (e: React.PointerEvent<HTMLCanvasElement>) => {
    if (!drawingRef.current) return;
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    const { x, y } = pointerPos(e, canvas);
    ctx.lineTo(x, y);
    ctx.stroke();
    if (!hasInk) setHasInk(true);
  };

  const end = (e: React.PointerEvent<HTMLCanvasElement>) => {
    drawingRef.current = false;
    canvasRef.current?.releasePointerCapture(e.pointerId);
  };

  const clear = () => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    setHasInk(false);
  };

  const save = async () => {
    const canvas = canvasRef.current;
    if (!canvas || !hasInk) return;
    const blob = await new Promise<Blob | null>((resolve) => {
      canvas.toBlob((b) => resolve(b), 'image/png');
    });
    if (!blob) return;
    const bytes = await blob.arrayBuffer();
    onCapture({ bytes, mime: 'image/png' });
    clear();
  };

  return (
    <div style={padWrapStyle} data-testid="drawn-signature-pad">
      <canvas
        ref={canvasRef}
        width={width}
        height={height}
        style={padCanvasStyle(width, height)}
        onPointerDown={start}
        onPointerMove={move}
        onPointerUp={end}
        onPointerLeave={end}
        data-testid="drawn-signature-canvas"
      />
      <div style={padActionsStyle}>
        <button type="button" onClick={clear} style={secondaryBtnStyle(false)}>
          {clearLabel}
        </button>
        <button
          type="button"
          onClick={save}
          disabled={!hasInk}
          style={primaryBtnStyle(!hasInk)}
          data-testid="drawn-signature-save"
        >
          {saveLabel}
        </button>
      </div>
    </div>
  );
}

function pointerPos(e: React.PointerEvent<HTMLCanvasElement>, canvas: HTMLCanvasElement) {
  const rect = canvas.getBoundingClientRect();
  const scaleX = canvas.width / rect.width;
  const scaleY = canvas.height / rect.height;
  return {
    x: (e.clientX - rect.left) * scaleX,
    y: (e.clientY - rect.top) * scaleY,
  };
}

// ---------------------------------------------------------------
// Typed
// ---------------------------------------------------------------

export interface TypedSignatureFieldProps {
  onCapture: (sig: CapturedSignature) => void;
  defaultText?: string;
  saveLabel?: string;
}

export function TypedSignatureField({
  onCapture,
  defaultText = '',
  saveLabel = 'Use this signature',
}: TypedSignatureFieldProps) {
  const [value, setValue] = useState(defaultText);
  const save = () => {
    const trimmed = value.trim();
    if (!trimmed) return;
    const bytes = new TextEncoder().encode(trimmed).buffer;
    onCapture({ bytes, mime: 'text/plain' });
    setValue('');
  };
  return (
    <div style={padWrapStyle} data-testid="typed-signature-field">
      <input
        type="text"
        value={value}
        onChange={(e) => setValue(e.target.value)}
        placeholder="Type your full name"
        style={typedInputStyle}
        data-testid="typed-signature-input"
        autoFocus
      />
      <div style={padActionsStyle}>
        <button
          type="button"
          onClick={save}
          disabled={!value.trim()}
          style={primaryBtnStyle(!value.trim())}
          data-testid="typed-signature-save"
        >
          {saveLabel}
        </button>
      </div>
    </div>
  );
}

// ---------------------------------------------------------------
// Uploaded
// ---------------------------------------------------------------

export interface UploadedSignatureFieldProps {
  onCapture: (sig: CapturedSignature) => void;
  /** Accept attribute. Default image/*. */
  accept?: string;
}

export function UploadedSignatureField({
  onCapture,
  accept = 'image/png,image/jpeg,image/svg+xml',
}: UploadedSignatureFieldProps) {
  const inputRef = useRef<HTMLInputElement>(null);
  const [fileName, setFileName] = useState<string | null>(null);
  const onChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const bytes = await file.arrayBuffer();
    onCapture({ bytes, mime: file.type || 'application/octet-stream' });
    setFileName(file.name);
    if (inputRef.current) inputRef.current.value = '';
  };
  return (
    <div style={padWrapStyle} data-testid="uploaded-signature-field">
      <label style={uploadLabelStyle}>
        <input
          ref={inputRef}
          type="file"
          accept={accept}
          onChange={onChange}
          style={{ display: 'none' }}
          data-testid="uploaded-signature-input"
        />
        <span>{fileName ?? 'Choose image…'}</span>
      </label>
    </div>
  );
}

// ---------------------------------------------------------------
// Shared styles
// ---------------------------------------------------------------

const padWrapStyle: CSSProperties = {
  display: 'flex',
  flexDirection: 'column',
  gap: 10,
};

const padCanvasStyle = (w: number, h: number): CSSProperties => ({
  width: w,
  height: h,
  maxWidth: '100%',
  border: '1px dashed var(--doc-border, #cbd5e1)',
  borderRadius: 8,
  background: 'var(--doc-surface, #fff)',
  cursor: 'crosshair',
  touchAction: 'none',
});

const padActionsStyle: CSSProperties = {
  display: 'flex',
  justifyContent: 'flex-end',
  gap: 8,
};

const typedInputStyle: CSSProperties = {
  width: '100%',
  padding: '10px 12px',
  border: '1px solid var(--doc-border, #cbd5e1)',
  borderRadius: 6,
  fontSize: 18,
  fontFamily: '"Caveat", "Dancing Script", "Brush Script MT", cursive',
  background: 'var(--doc-surface, #fff)',
  color: 'var(--doc-text, #0f172a)',
};

const uploadLabelStyle: CSSProperties = {
  display: 'inline-flex',
  padding: '8px 14px',
  border: '1px dashed var(--doc-border, #cbd5e1)',
  borderRadius: 6,
  background: 'var(--doc-surface, #fff)',
  color: 'var(--doc-text, #0f172a)',
  fontSize: 13,
  cursor: 'pointer',
  alignSelf: 'flex-start',
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

function secondaryBtnStyle(disabled: boolean): CSSProperties {
  return {
    padding: '8px 16px',
    borderRadius: 6,
    border: '1px solid var(--doc-border, #cbd5e1)',
    background: 'transparent',
    color: 'var(--doc-text, #0f172a)',
    fontSize: 13,
    fontWeight: 500,
    cursor: disabled ? 'not-allowed' : 'pointer',
    fontFamily: 'inherit',
  };
}
