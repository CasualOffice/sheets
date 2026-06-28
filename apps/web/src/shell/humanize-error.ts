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
 * Map common open-file / parse / save error messages into human-friendly
 * sentences. Used by the LoadingOverlay error card and the autosave
 * banner so users don't see raw ExcelJS / JSZip / parser stack traces.
 *
 * The classifier is intentionally a flat string-match list, not a
 * regex tree — every entry is one observed failure mode pulled from
 * real e2e or user-report sessions. When a new failure shape comes
 * in, add a new rule rather than refactoring; the cost of one extra
 * `if` is nothing next to the cost of a user reading "ENOENT" and
 * filing a confused issue.
 *
 * The raw error is preserved by the caller in a collapsible
 * "Technical details" block — power users + bug reporters still get
 * everything they need.
 */

export type HumanizedError = {
  /** Short, scannable headline. ≤ 60 chars. Renders as a sentence. */
  title: string;
  /** Optional one-line explainer rendered under the headline. */
  hint?: string;
  /** Suggested next action label — feeds the retry button text. */
  retryLabel?: string;
};

/**
 * Translate a raw error message into a friendly headline + optional
 * hint. `fileName` is only used to personalize the headline; the
 * caller decides what to do with the original error string.
 */
export function humanizeOpenError(rawMessage: string, fileName: string): HumanizedError {
  const m = rawMessage.toLowerCase();

  // ── ZIP / file-format failures — ExcelJS unwraps .xlsx as a zip;
  //    a corrupt or wrong-format file lands here.
  if (
    m.includes('end of central directory') ||
    m.includes('central directory') ||
    m.includes('not a zip') ||
    m.includes('invalid zip') ||
    m.includes('corrupted zip')
  ) {
    return {
      title: `${fileName} doesn't look like a valid spreadsheet file.`,
      hint: 'The file may be damaged, partially downloaded, or not a real .xlsx / .ods.',
      retryLabel: 'Pick a different file',
    };
  }

  // ── Unsupported file shape (encrypted .xlsx, password-protected,
  //    or an .xls binary not covered by ExcelJS).
  if (
    m.includes('encrypted') ||
    m.includes('password') ||
    m.includes('protected') ||
    m.includes('biff') ||
    m.includes('legacy .xls')
  ) {
    return {
      title: `${fileName} is encrypted or in an unsupported format.`,
      hint: 'Casual Sheets opens unprotected .xlsx, .ods, .csv and .tsv files. Save a copy without the password, then try again.',
      retryLabel: 'Open a different file',
    };
  }

  // ── Network / fetch failures (template download, share-seed fetch).
  if (
    m.includes('failed to fetch') ||
    m.includes('networkerror') ||
    m.includes('network error') ||
    m.includes('econnrefused') ||
    m.includes('econnreset') ||
    m.includes('aborterror') ||
    m.includes('aborted')
  ) {
    return {
      title: `Couldn't reach the server while opening ${fileName}.`,
      hint: 'Check your connection and try again. If the demo is offline this will resolve once it comes back.',
      retryLabel: 'Try again',
    };
  }

  // ── HTTP 4xx / 5xx envelope errors (the seed download paths surface
  //    `HTTP 404` / `HTTP 500` from the gateway).
  if (m.includes('http 404') || m.includes('not found')) {
    return {
      title: `${fileName} is no longer available.`,
      hint: 'The share link may have expired or been deleted by its owner.',
    };
  }
  if (m.includes('http 403') || m.includes('forbidden')) {
    return {
      title: `You don't have access to ${fileName}.`,
      hint: 'Ask the owner for a fresh share link with edit or view permission.',
    };
  }
  if (m.includes('http 5') || m.includes('internal server error') || m.includes('bad gateway')) {
    return {
      title: `Server hiccup while opening ${fileName}.`,
      hint: 'This is usually transient — give it a few seconds and try again.',
      retryLabel: 'Try again',
    };
  }

  // ── ODS-specific path complaints (the @casualoffice/core ods loader can
  //    throw "expected mimetype application/vnd.oasis…" for files
  //    masquerading as .ods).
  if (m.includes('mimetype') || m.includes('content.xml') || m.includes('ods')) {
    return {
      title: `${fileName} doesn't look like a valid .ods file.`,
      hint: 'Try re-saving from your spreadsheet app, or convert to .xlsx first.',
      retryLabel: 'Pick a different file',
    };
  }

  // ── Memory / size envelope — really big files can hit Chrome's
  //    256 MB heap on a tab. We try to detect via the message.
  if (
    m.includes('out of memory') ||
    m.includes('allocation failed') ||
    m.includes('rangeerror') ||
    m.includes('maximum call stack')
  ) {
    return {
      title: `${fileName} is too large for the browser to load.`,
      hint: 'Try a smaller workbook, or open it in desktop Excel and copy the sheet you need.',
    };
  }

  // ── Anything else — give a friendly headline + tuck the raw text
  //    away. The caller renders the raw message in a collapsible
  //    "Technical details" section for diagnostics.
  return {
    title: `Couldn't open ${fileName}.`,
    hint: 'The file may be damaged, the format unsupported, or the connection dropped. Try again or pick a different file.',
    retryLabel: 'Try again',
  };
}
