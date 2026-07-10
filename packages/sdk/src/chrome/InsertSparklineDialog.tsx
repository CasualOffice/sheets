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
 * InsertSparklineDialog — the SDK chrome's built-in "Insert sparkline" modal.
 *
 * Excel/Sheets-style in-cell sparklines: pick a type (line / column / win-loss),
 * a data range (the values to plot), and a location cell (where the mini-chart
 * renders). Follows the FormatCells/DataValidation exemplars: reads the active
 * A1 selection off the FUniver facade to seed the fields, gathers the form, and
 * applies through the SDK API.
 *
 * FACADE NOTE — no first-class Univer sparkline facade is installed. Unlike
 * data-validation (`@univerjs/sheets-data-validation/facade`), there is no
 * `@univerjs/sheets-sparkline` package in `node_modules/@univerjs/*`, so there is
 * no `univerAPI.newSparkline()` builder or `FRange.setSparkline(...)` to call.
 * Sparklines in this workbook are a Casual-Sheets feature persisted on
 * `IWorkbookData.resources['__casual_sheets_sparklines__']` — the resource the
 * app's `SparklineLayer` renders from and that round-trips through xlsx/collab
 * (see apps/web/src/sparklines/{types,resources}.ts). The dialog therefore wires
 * the REAL op the feature owns: it merges a new `SparklineModel` into that
 * resource on the live snapshot via `api.getContent()` and re-mounts it via
 * `api.setContent()` (the canonical cross-editor content accessors on
 * `CasualSheetsAPI`). The resource name + model shape are duplicated locally so
 * the SDK stays decoupled from the app package.
 *
 * Mounted by `<DialogHost>` when `openDialog('insert-sparkline')` is called and
 * no host override is registered.
 */

import { useMemo, useState, type CSSProperties } from 'react';
import type { IWorkbookData } from '@univerjs/core';
import type { DialogComponentProps } from './extensions';
import type { CasualSheetsAPI } from '../sheets/api';
import { Dialog } from './Dialog';
import {
  DIALOG_BTN_PRIMARY_STYLE,
  DIALOG_BTN_SECONDARY_STYLE,
  DIALOG_FIELD_STYLE,
  DIALOG_INPUT_STYLE,
  DIALOG_LABEL_STYLE,
} from './dialog-styles';

/** Excel-canonical sparkline families. Mirrors apps/web sparklines `SparklineType`. */
type SparklineType = 'line' | 'column' | 'win-loss';

/** A single source rectangle in row/column indices. */
interface SourceRange {
  startRow: number;
  endRow: number;
  startColumn: number;
  endColumn: number;
}

/**
 * Sparkline model persisted on the workbook resource. Structurally identical to
 * the app's `SparklineModel` (apps/web/src/sparklines/types.ts) — duplicated here
 * so the SDK doesn't depend on the app package. `SparklineLayer` reads this shape.
 */
interface SparklineModel {
  id: string;
  type: SparklineType;
  unitId: string;
  sheetId: string;
  source: SourceRange;
  anchor: { row: number; col: number };
}

/** Resource key + payload envelope — mirrors `SPARKLINES_RESOURCE_NAME` /
 *  `SparklinesResourceV1` in apps/web/src/sparklines/types.ts. */
const SPARKLINES_RESOURCE_NAME = '__casual_sheets_sparklines__';
interface SparklinesResourceV1 {
  v: 1;
  sparklines: SparklineModel[];
}

const TYPE_OPTIONS: Array<{ value: SparklineType; label: string }> = [
  { value: 'line', label: 'Line' },
  { value: 'column', label: 'Column' },
  { value: 'win-loss', label: 'Win / Loss' },
];

interface DialogState {
  type: SparklineType;
  /** Data range in A1, e.g. "A1:F1". */
  sourceA1: string;
  /** Location cell in A1, e.g. "G1". */
  anchorA1: string;
}

/** The active FRange, or null when there is no selection. */
function activeRange(api: CasualSheetsAPI) {
  return api.univer.getActiveWorkbook()?.getActiveSheet()?.getActiveRange() ?? null;
}

/** Convert column letters (A, B, …, AA) to a 0-based index. */
function colLettersToIndex(letters: string): number {
  let n = 0;
  for (let i = 0; i < letters.length; i += 1) {
    n = n * 26 + (letters.charCodeAt(i) - 64);
  }
  return n - 1;
}

/** Parse an A1 range ("A1:F1" or a single "A1") into row/column indices. */
function parseRange(s: string): SourceRange | null {
  const m = /^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/.exec(s.trim().toUpperCase());
  if (!m) return null;
  const c1 = colLettersToIndex(m[1]);
  const r1 = parseInt(m[2], 10) - 1;
  const c2 = m[3] ? colLettersToIndex(m[3]) : c1;
  const r2 = m[4] ? parseInt(m[4], 10) - 1 : r1;
  return {
    startRow: Math.min(r1, r2),
    endRow: Math.max(r1, r2),
    startColumn: Math.min(c1, c2),
    endColumn: Math.max(c1, c2),
  };
}

/** Parse a single A1 cell ("G1") into row/col indices. */
function parseSingleCell(s: string): { row: number; col: number } | null {
  const m = /^\$?([A-Z]+)\$?(\d+)$/.exec(s.trim().toUpperCase());
  if (!m) return null;
  return { col: colLettersToIndex(m[1]), row: parseInt(m[2], 10) - 1 };
}

/**
 * Merge a new sparkline into the workbook's sparklines resource and re-mount the
 * snapshot. Returns false when there is no active workbook.
 *
 * This is the REAL persistence path for sparklines in this workbook (there is no
 * Univer sparkline facade — see the file header). `getContent()` yields the live
 * `IWorkbookData` including `resources`; we splice our model into the
 * `__casual_sheets_sparklines__` entry and hand it back through `setContent()`,
 * exactly the shape `SparklineLayer` renders and xlsx/collab round-trip.
 */
function insertSparkline(api: CasualSheetsAPI, model: SparklineModel): boolean {
  const data = api.getContent();
  if (!data) return false;

  const resources = data.resources ? [...data.resources] : [];
  const idx = resources.findIndex((r) => r.name === SPARKLINES_RESOURCE_NAME);

  let existing: SparklineModel[] = [];
  if (idx >= 0 && resources[idx]?.data) {
    try {
      const parsed = JSON.parse(resources[idx].data) as Partial<SparklinesResourceV1>;
      if (parsed?.v === 1 && Array.isArray(parsed.sparklines)) existing = parsed.sparklines;
    } catch {
      existing = [];
    }
  }

  const payload: SparklinesResourceV1 = { v: 1, sparklines: [...existing, model] };
  const entry = { name: SPARKLINES_RESOURCE_NAME, data: JSON.stringify(payload) };
  if (idx >= 0) resources[idx] = entry;
  else resources.push(entry);

  const next: IWorkbookData = { ...data, resources };
  api.setContent(next);
  return true;
}

/** Read the active workbook's unitId + active sheetId for the model. */
function activeIds(api: CasualSheetsAPI): { unitId: string; sheetId: string } | null {
  const wb = api.univer.getActiveWorkbook();
  const sheet = wb?.getActiveSheet();
  if (!wb || !sheet) return null;
  return { unitId: wb.getId(), sheetId: sheet.getSheetId() };
}

const RANGE_NOTE_STYLE: CSSProperties = {
  fontSize: 12,
  color: 'var(--cs-chrome-muted, #605e5c)',
  marginBottom: 12,
};

const RADIO_ROW_STYLE: CSSProperties = {
  display: 'flex',
  gap: 16,
  marginTop: 2,
};

const RADIO_OPT_STYLE: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 6,
  cursor: 'pointer',
};

const ERROR_STYLE: CSSProperties = {
  fontSize: 12,
  color: 'var(--cs-chrome-danger-fg, #b91c1c)',
  marginTop: 4,
};

export function InsertSparklineDialog({ api, onClose }: DialogComponentProps) {
  // Seed the data range from the current selection (its A1 label), leaving the
  // location cell for the user to pick — mirrors Excel's Insert Sparkline flow.
  const selectedA1 = useMemo(() => {
    const fRange = activeRange(api) as unknown as { getA1Notation?: () => string } | null;
    return fRange?.getA1Notation?.() ?? '';
  }, [api]);

  const [state, setState] = useState<DialogState>({
    type: 'line',
    sourceA1: selectedA1,
    anchorA1: '',
  });
  const [error, setError] = useState<string | null>(null);

  const hasWorkbook = api.univer.getActiveWorkbook() != null;

  const update = <K extends keyof DialogState>(key: K, value: DialogState[K]) => {
    setState((prev) => ({ ...prev, [key]: value }));
    setError(null);
  };

  const apply = () => {
    const source = parseRange(state.sourceA1);
    if (!source) {
      setError('Data range must be a range like A1:F1.');
      return;
    }
    const anchor = parseSingleCell(state.anchorA1);
    if (!anchor) {
      setError('Location must be a single cell like G1.');
      return;
    }
    const ids = activeIds(api);
    if (!ids) {
      setError('No active sheet.');
      return;
    }

    const model: SparklineModel = {
      id: `spark-${Math.random().toString(36).slice(2, 10)}`,
      type: state.type,
      unitId: ids.unitId,
      sheetId: ids.sheetId,
      source,
      anchor,
    };
    if (insertSparkline(api, model)) onClose();
    else setError('Could not read the active workbook.');
  };

  return (
    <Dialog
      title="Insert sparkline"
      onClose={onClose}
      width={420}
      data-testid="cs-insert-sparkline-dialog"
      footer={
        <>
          <button type="button" style={DIALOG_BTN_SECONDARY_STYLE} onClick={onClose}>
            Cancel
          </button>
          <button
            type="button"
            style={DIALOG_BTN_PRIMARY_STYLE}
            data-testid="cs-insert-sparkline-apply"
            disabled={!hasWorkbook}
            onClick={apply}
          >
            Insert
          </button>
        </>
      }
    >
      <div style={RANGE_NOTE_STYLE}>
        Draws an in-cell mini-chart in the location cell from the values in the data range.
      </div>

      <label style={DIALOG_FIELD_STYLE}>
        <span style={DIALOG_LABEL_STYLE}>Type</span>
        <div
          style={RADIO_ROW_STYLE}
          role="radiogroup"
          aria-label="Sparkline type"
          data-testid="cs-insert-sparkline-type"
        >
          {TYPE_OPTIONS.map((opt) => (
            <label key={opt.value} style={RADIO_OPT_STYLE}>
              <input
                type="radio"
                name="cs-sparkline-type"
                value={opt.value}
                data-testid={`cs-insert-sparkline-type-${opt.value}`}
                checked={state.type === opt.value}
                onChange={() => update('type', opt.value)}
              />
              <span>{opt.label}</span>
            </label>
          ))}
        </div>
      </label>

      <label style={DIALOG_FIELD_STYLE}>
        <span style={DIALOG_LABEL_STYLE}>Data range</span>
        <input
          style={DIALOG_INPUT_STYLE}
          data-testid="cs-insert-sparkline-source"
          value={state.sourceA1}
          placeholder="A1:F1"
          spellCheck={false}
          onChange={(e) => update('sourceA1', e.target.value.toUpperCase())}
        />
      </label>

      <label style={DIALOG_FIELD_STYLE}>
        <span style={DIALOG_LABEL_STYLE}>Location</span>
        <input
          style={DIALOG_INPUT_STYLE}
          data-testid="cs-insert-sparkline-anchor"
          value={state.anchorA1}
          placeholder="G1"
          spellCheck={false}
          onChange={(e) => update('anchorA1', e.target.value.toUpperCase())}
        />
      </label>

      {error && (
        <div style={ERROR_STYLE} data-testid="cs-insert-sparkline-error">
          {error}
        </div>
      )}
    </Dialog>
  );
}
