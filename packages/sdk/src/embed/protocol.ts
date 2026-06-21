/**
 * Iframe protocol — wire envelopes that match
 * `docs/internal/13-iframe-protocol.md`.
 *
 * Mirror updates to the doc whenever a new envelope shape lands.
 * The discriminator on `type` (always starts with `casual.`) and
 * the per-envelope `data` shape are the contract.
 */

import type {
  CancelReason,
  SignatureCompletePayload,
  SignatureField,
  SignatureMode,
  SignedFieldPayload,
} from '../signing/types';

export type CasualApp = 'docs' | 'sheet';

/** Common envelope shape — every postMessage on the wire matches this. */
export interface CasualEnvelope<T = unknown> {
  type: string;
  app: CasualApp;
  /** Per-request id for request/response correlation. Empty for fire-and-forget. */
  id?: string;
  /** Protocol version. Bumped only on breaking changes. */
  v: 1;
  data: T;
}

// ---------------------------------------------------------------
// Handshake
// ---------------------------------------------------------------

export interface EditorHelloData {
  capabilities: string[];
  version: string;
  commit: string;
}

export interface HostHelloData {
  capabilities: string[];
  authToken?: string;
}

// ---------------------------------------------------------------
// Load + save (editor → host requests; host → editor responses)
// ---------------------------------------------------------------

export interface LoadRequestData {
  docId: string;
}

export interface LoadResponseDataOk {
  ok: true;
  bytes: ArrayBuffer;
  etag?: string;
  fileName: string;
  readOnly?: boolean;
}

export interface LoadResponseDataErr {
  ok: false;
  code: string;
  message?: string;
}

export type LoadResponseData = LoadResponseDataOk | LoadResponseDataErr;

export interface SaveRequestData {
  docId: string;
  bytes: ArrayBuffer;
  baseEtag?: string;
}

export interface SaveResponseDataOk {
  ok: true;
  etag: string;
}

export interface SaveResponseDataErr {
  ok: false;
  code: string;
  etag?: string;
  message?: string;
}

export type SaveResponseData = SaveResponseDataOk | SaveResponseDataErr;

// ---------------------------------------------------------------
// Save / exit notifications (editor → host, fire-and-forget)
// ---------------------------------------------------------------
//
// The lightweight counterpart to the bytes-carrying load/save *request*
// pair above. These mirror the SDK's React `onSave` / `onExit` hooks one
// for one — the "one shape, two surfaces" save/exit contract — so a host
// that frames the iframe gets the same persistence signals a host that
// renders `<CasualSheets>` directly does. Fire-and-forget: the editor
// never owns storage; the host decides what to do with the snapshot.
//
// `casual.save.request` (above) stays the WOPI-style path for hosts that
// want xlsx bytes + etag round-trips; these notifications are the simpler
// "here's the current state, persist it however you like" path.

/** Editor → host: the user explicitly asked to save — Ctrl/Cmd+S inside
 *  the iframe, or the host's `casual.command.save`. Carries the full
 *  editor snapshot as JSON (sheet app: Univer's `IWorkbookData`). Mirror
 *  of the React `onSave` hook. v0.9+. */
export interface SaveNotifyData {
  /** App-specific snapshot JSON. Sheet: `IWorkbookData`. */
  snapshot: unknown;
  /** What triggered the save: the in-editor shortcut or a host command. */
  reason: 'shortcut' | 'host';
}

/** Editor → host: the editor is unmounting / navigating away. Carries
 *  the final snapshot so the host can persist on exit. Mirror of the
 *  React `onExit` hook. v0.9+. */
export interface ExitData {
  /** App-specific snapshot JSON. Sheet: `IWorkbookData`. */
  snapshot: unknown;
}

// ---------------------------------------------------------------
// Selection + telemetry + lock (editor → host notifications)
// ---------------------------------------------------------------

export interface SelectionChangedData {
  docs?: { paraId: string; from: number; to: number; selectedText: string };
  sheet?: { sheet: string; from: string; to: string };
}

export interface TelemetryEventData {
  kind: string;
  /** Arbitrary metric fields. */
  [k: string]: unknown;
}

export interface LockLostData {
  reason: 'taken_by_other' | 'expired' | 'host_revoked';
}

// ---------------------------------------------------------------
// Commands (host → editor)
// ---------------------------------------------------------------

export interface CommandSetReadOnlyData {
  readOnly: boolean;
}

export interface CommandSetThemeData {
  theme: 'light' | 'dark' | 'system';
}

/** Host → editor: enable/disable chrome features. Each key maps a toolbar
 *  group / menu item / capability to a boolean; a `false` hides the control
 *  AND blocks its command. Omitted keys default to enabled. Mirrors the
 *  `features` prop on `<CasualSheets>`. */
export interface CommandSetFeaturesData {
  features: Record<string, boolean>;
}

/** Editor → host: a chrome control backed by a dialog the SDK doesn't render
 *  itself (Format Cells, Insert Chart, Find & Replace, …) was activated. The
 *  host renders its OWN dialog/popup and applies the result via
 *  `executeCommand`. Fired only when the host opted into host-owned dialogs
 *  by handling this event. */
export interface DialogRequestData {
  kind: string;
  context?: unknown;
}

export interface CommandSetLocaleData {
  locale: string;
}

/** Host → editor: switch chrome density between the two consumer-facing
 *  modes without re-mounting. `preview` hides toolbar / formula bar /
 *  side panel / status bar / sheet tabs and runs read-only; `editor`
 *  shows the full UI. Mirrors the `viewMode` prop on `<CasualSheetsIframe>`. */
export interface CommandSetViewModeData {
  viewMode: 'preview' | 'editor';
}

/** Host → editor: execute a formatting / navigation command against
 *  the active selection in the embedded workbook. Hosts (like Casual
 *  Drive) build their own toolbar above the iframe and dispatch these
 *  commands instead of Univer's built-in ribbon, which the SDK can't
 *  ship because the ribbon resolves IRPCChannelService at construction
 *  and that service needs a worker the SDK doesn't bundle.
 *
 *  v0.6 covered the toggle set; v0.7 adds the rich-format set (font,
 *  size, colour, fill) + merge / unmerge. Arg-carrying commands read
 *  the relevant field off `args` — every other command ignores it. */
export interface CommandExecuteData {
  command: // v0.6 — toggle / nav (no args)
    | 'undo'
    | 'redo'
    | 'bold'
    | 'italic'
    | 'underline'
    | 'strikethrough'
    | 'align-left'
    | 'align-center'
    | 'align-right'
    // v0.7 — rich format (args carry the value)
    | 'set-font-family'
    | 'set-font-size'
    | 'set-text-color'
    | 'reset-text-color'
    | 'set-bg-color'
    | 'reset-bg-color'
    | 'merge'
    | 'unmerge'
    // v0.8 — number formats + freeze + wrap
    | 'numfmt-currency'
    | 'numfmt-percent'
    | 'numfmt-add-decimal'
    | 'numfmt-subtract-decimal'
    | 'numfmt-custom'
    | 'wrap-toggle'
    | 'freeze-first-row'
    | 'freeze-first-column'
    | 'freeze-none';
  args?: {
    /** Used by `set-font-family`. */
    family?: string;
    /** Used by `set-font-size`. Integer point size. */
    size?: number;
    /** Used by `set-text-color` and `set-bg-color`. Hex like `#1a73e8`. */
    color?: string;
    /** Used by `numfmt-custom`. The Excel-style format string, e.g.
     *  `"#,##0.00"`, `"$#,##0"`, `"0.00%"`, `"d-mmm-yy"`. v0.8+. */
    pattern?: string;
  };
}

/** Editor → host: emitted whenever the selection's active cell's
 *  format flags change. Drive's toolbar mirrors this state in the
 *  button "pressed" / value indicators so the surface always reflects
 *  what the user would see in the cell. v0.7 widens the payload with
 *  the rich-format read-back (fontFamily, fontSize, textColor, bgColor)
 *  so the toolbar's font picker / size stepper / colour swatches stay
 *  in sync without the host having to poll. */
export interface SelectionFormatStateData {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strikethrough: boolean;
  align: 'left' | 'center' | 'right' | null;
  /** Defined font family on the active cell, or null when the cell
   *  inherits the workbook default. v0.7+. */
  fontFamily: string | null;
  /** Defined font size on the active cell, or null when the cell
   *  inherits the workbook default. v0.7+. */
  fontSize: number | null;
  /** Hex text colour like `#1a73e8`, or null when default. v0.7+. */
  textColor: string | null;
  /** Hex background colour, or null when no fill is set. v0.7+. */
  bgColor: string | null;
}

// ---------------------------------------------------------------
// Errors (editor → host fatal signals)
// ---------------------------------------------------------------

/** Editor → host: a fatal error during boot / load. Hosts surface this
 *  via the wrapper's `onError` callback. */
export interface CasualErrorData {
  code: 'embed_not_served' | 'load_failed' | 'parse_failed' | 'boot_failed' | 'internal';
  message: string;
}

// ---------------------------------------------------------------
// Signing (uniform with the SDK `signing` prop)
// ---------------------------------------------------------------

export interface SignatureRequestData {
  fields: SignatureField[];
  mode: SignatureMode;
  banner?: string;
}

export interface SignatureRequestAckData {
  ok: boolean;
  code?: string;
}

export type SignatureFieldSignedData = SignedFieldPayload;
export type SignatureCompleteData = SignatureCompletePayload;

export interface SignatureCancelData {
  reason: CancelReason;
}

// ---------------------------------------------------------------
// Type guards
// ---------------------------------------------------------------

export function isCasualEnvelope(value: unknown): value is CasualEnvelope {
  if (!value || typeof value !== 'object') return false;
  const v = value as Record<string, unknown>;
  return (
    typeof v.type === 'string' &&
    v.type.startsWith('casual.') &&
    (v.app === 'docs' || v.app === 'sheet') &&
    v.v === 1 &&
    'data' in v
  );
}
