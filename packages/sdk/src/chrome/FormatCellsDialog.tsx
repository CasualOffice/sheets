/**
 * FormatCellsDialog — the SDK chrome's built-in Format Cells modal.
 *
 * Ported from `apps/web/src/shell/FormatCellsDialog.tsx`, reimplemented against
 * the `CasualSheetsAPI` + FUniver facade ONLY (no app context: no use-univer, no
 * useActiveCellState, no home-tab-actions). Tabs: Number · Alignment · Font ·
 * Border · Fill. Reads the active cell's current style off the snapshot to seed
 * the form, and applies via the same facade calls the app uses
 * (`setFontWeight` / `setHorizontalAlignment` / `setNumberFormat` / `setBorder`
 * / `setBackground` / …), so behaviour matches the real app exactly.
 *
 * Mounted by `<DialogHost>` when `openDialog('format-cells')` is called and no
 * host override is registered.
 */

import { useEffect, useMemo, useState, type CSSProperties } from 'react';
import { BorderStyleTypes, BorderType } from '@univerjs/core';
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

type TabId = 'number' | 'alignment' | 'font' | 'border' | 'fill';

type HAlign = 'left' | 'center' | 'right';
type VAlign = 'top' | 'middle' | 'bottom';
type BorderChoice = 'all' | 'outside' | 'top' | 'bottom' | 'left' | 'right' | 'none';

const DEFAULT_BORDER_COLOR = '#000000';

/** Number-format presets — patterns copied verbatim from the app's
 *  `home-tab-actions.NUMBER_FORMAT_PATTERNS`. */
const NUMBER_FORMAT_PATTERNS = {
  general: '',
  number: '#,##0.00',
  integer: '#,##0',
  currency: '"$"#,##0.00',
  accounting: '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)',
  percent: '0.00%',
  date: 'yyyy-mm-dd',
  time: 'hh:mm:ss',
  scientific: '0.00E+00',
  text: '@',
} as const;
type NumberFormatKey = keyof typeof NUMBER_FORMAT_PATTERNS;

const WEIGHT_TO_STYLE: BorderStyleTypes = BorderStyleTypes.THIN;

interface DialogState {
  numberFormatKey: NumberFormatKey;
  numberPattern: string;
  align: HAlign;
  vAlign: VAlign;
  wrap: boolean;
  fontFamily: string;
  fontSize: string;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  borderChoice: BorderChoice;
  borderColor: string;
  fillColor: string;
}

const TABS: Array<{ id: TabId; label: string }> = [
  { id: 'number', label: 'Number' },
  { id: 'alignment', label: 'Alignment' },
  { id: 'font', label: 'Font' },
  { id: 'border', label: 'Border' },
  { id: 'fill', label: 'Fill' },
];

const FONT_FAMILIES = [
  'Calibri',
  'Arial',
  'Helvetica',
  'Inter',
  'Times New Roman',
  'Georgia',
  'Verdana',
  'Courier New',
  'JetBrains Mono',
];

const NUMBER_FORMAT_OPTIONS: Array<{ value: NumberFormatKey; label: string }> = [
  { value: 'general', label: 'General' },
  { value: 'integer', label: 'Number (no decimals)' },
  { value: 'number', label: 'Number (2 decimals)' },
  { value: 'currency', label: 'Currency' },
  { value: 'accounting', label: 'Accounting' },
  { value: 'percent', label: 'Percent' },
  { value: 'date', label: 'Date' },
  { value: 'time', label: 'Time' },
  { value: 'scientific', label: 'Scientific' },
  { value: 'text', label: 'Text' },
];

function detectFormatKey(pattern: string): NumberFormatKey {
  for (const key of Object.keys(NUMBER_FORMAT_PATTERNS) as NumberFormatKey[]) {
    if (NUMBER_FORMAT_PATTERNS[key] === pattern) return key;
  }
  return 'general';
}

/** Univer style shape we read off the active cell (subset). */
interface CellStyle {
  ff?: string;
  fs?: number;
  bl?: number; // bold (1)
  it?: number; // italic (1)
  ul?: { s?: number } | number; // underline
  ht?: number; // horizontal: 1 left, 2 center, 3 right
  vt?: number; // vertical: 1 top, 2 middle, 3 bottom
  tb?: number; // text wrap: 3 wrap
  bg?: { rgb?: string };
  n?: { pattern?: string };
  bd?: Record<string, unknown>;
}

/**
 * Resolve the active cell's style from the snapshot. Styles can be an inline
 * object or a string key into the workbook's style table — mirror the app's
 * `getCellData()` + styles-table lookup, but off the SDK snapshot.
 */
function readActiveStyle(api: CasualSheetsAPI): CellStyle | null {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const snap = api.getSnapshot() as any;
  if (!snap) return null;
  const sel = api.getSelection();
  const sheetId = sel?.sheetId ?? api.univer.getActiveWorkbook()?.getActiveSheet()?.getSheetId();
  if (!sheetId) return null;
  const sheet = snap.sheets?.[sheetId];
  if (!sheet) return null;
  const row = sel?.range.startRow ?? 0;
  const col = sel?.range.startColumn ?? 0;
  const cell = sheet.cellData?.[row]?.[col];
  if (!cell) return null;
  const s = cell.s;
  if (typeof s === 'string') return (snap.styles?.[s] as CellStyle) ?? null;
  return (s as CellStyle) ?? null;
}

function buildInitialState(style: CellStyle | null): DialogState {
  const pattern = style?.n?.pattern ?? '';
  const underline =
    typeof style?.ul === 'number' ? style.ul === 1 : !!(style?.ul && style.ul.s === 1);
  return {
    numberFormatKey: detectFormatKey(pattern),
    numberPattern: pattern,
    align: style?.ht === 2 ? 'center' : style?.ht === 3 ? 'right' : 'left',
    vAlign: style?.vt === 1 ? 'top' : style?.vt === 2 ? 'middle' : 'bottom',
    wrap: style?.tb === 3,
    fontFamily: style?.ff || 'Calibri',
    fontSize: String(style?.fs || 11),
    bold: style?.bl === 1,
    italic: style?.it === 1,
    underline,
    borderChoice: style?.bd && Object.keys(style.bd).length > 0 ? 'outside' : 'none',
    borderColor: DEFAULT_BORDER_COLOR,
    fillColor: style?.bg?.rgb || '#ffffff',
  };
}

function activeRange(api: CasualSheetsAPI) {
  return api.univer.getActiveWorkbook()?.getActiveSheet()?.getActiveRange() ?? null;
}

/** Apply every section of the form via the facade — same calls as the app. */
function applyFormat(api: CasualSheetsAPI, s: DialogState): void {
  const range = activeRange(api);
  if (!range) return;

  // Number format — sheets-numfmt facade extension (runtime-added to FRange).
  (range as unknown as { setNumberFormat?: (p: string) => unknown }).setNumberFormat?.(
    s.numberPattern,
  );

  // Univer's facade uses 'normal' for right-aligned (FHorizontalAlignment).
  range.setHorizontalAlignment(s.align === 'right' ? 'normal' : s.align);
  range.setVerticalAlignment(s.vAlign);
  range.setWrap(s.wrap);

  range.setFontFamily(s.fontFamily || null);
  const size = Number(s.fontSize);
  if (Number.isFinite(size) && size > 0) range.setFontSize(size);
  range.setFontWeight(s.bold ? 'bold' : 'normal');
  range.setFontStyle(s.italic ? 'italic' : 'normal');
  range.setFontLine(s.underline ? 'underline' : 'none');

  // Borders — map the preset to a Univer BorderType + thin style (matches the
  // app's setBorders default weight).
  const type =
    s.borderChoice === 'all'
      ? BorderType.ALL
      : s.borderChoice === 'outside'
        ? BorderType.OUTSIDE
        : s.borderChoice === 'top'
          ? BorderType.TOP
          : s.borderChoice === 'bottom'
            ? BorderType.BOTTOM
            : s.borderChoice === 'left'
              ? BorderType.LEFT
              : s.borderChoice === 'right'
                ? BorderType.RIGHT
                : BorderType.NONE;
  const borderStyle = s.borderChoice === 'none' ? BorderStyleTypes.NONE : WEIGHT_TO_STYLE;
  range.setBorder(type, borderStyle, s.borderColor);

  range.setBackground(s.fillColor);
}

const TABS_ROW_STYLE: CSSProperties = {
  display: 'flex',
  gap: 2,
  borderBottom: '1px solid var(--cs-chrome-border, #edeff3)',
  marginBottom: 14,
};

function tabStyle(activeTab: boolean): CSSProperties {
  return {
    padding: '6px 12px',
    border: 'none',
    borderBottom: `2px solid ${activeTab ? 'var(--cs-chrome-active-fg, #0e7490)' : 'transparent'}`,
    background: 'transparent',
    color: activeTab ? 'var(--cs-chrome-active-fg, #0e7490)' : 'var(--cs-chrome-muted, #605e5c)',
    font: 'inherit',
    fontSize: 13,
    fontWeight: activeTab ? 600 : 400,
    cursor: 'pointer',
  };
}

const CHECK_STYLE: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 6,
  marginBottom: 8,
  cursor: 'pointer',
};

const COLOR_INPUT_STYLE: CSSProperties = {
  width: 48,
  height: 30,
  padding: 2,
  border: '1px solid var(--cs-chrome-border, #cdd3db)',
  borderRadius: 6,
  background: 'var(--cs-chrome-input-bg, #fff)',
  cursor: 'pointer',
};

export function FormatCellsDialog({ api, onClose }: DialogComponentProps) {
  const [tab, setTab] = useState<TabId>('number');
  // Seed once from the active cell when the dialog mounts.
  const initial = useMemo(() => buildInitialState(readActiveStyle(api)), [api]);
  const [state, setState] = useState<DialogState>(initial);

  // Re-seed if the active cell changes while open (rare, but keeps it honest).
  useEffect(() => {
    setState(buildInitialState(readActiveStyle(api)));
  }, [api]);

  const update = <K extends keyof DialogState>(key: K, value: DialogState[K]) =>
    setState((prev) => ({ ...prev, [key]: value }));

  const apply = () => {
    applyFormat(api, state);
    onClose();
  };

  return (
    <Dialog
      title="Format Cells"
      onClose={onClose}
      width={480}
      data-testid="cs-format-cells-dialog"
      footer={
        <>
          <button type="button" style={DIALOG_BTN_SECONDARY_STYLE} onClick={onClose}>
            Cancel
          </button>
          <button
            type="button"
            style={DIALOG_BTN_PRIMARY_STYLE}
            data-testid="cs-format-cells-apply"
            onClick={apply}
          >
            Apply
          </button>
        </>
      }
    >
      <div style={TABS_ROW_STYLE} role="tablist" aria-label="Format sections">
        {TABS.map((t) => (
          <button
            key={t.id}
            type="button"
            role="tab"
            aria-selected={tab === t.id}
            data-testid={`cs-format-cells-tab-${t.id}`}
            style={tabStyle(tab === t.id)}
            onClick={() => setTab(t.id)}
          >
            {t.label}
          </button>
        ))}
      </div>

      <div role="tabpanel">
        {tab === 'number' && (
          <>
            <label style={DIALOG_FIELD_STYLE}>
              <span style={DIALOG_LABEL_STYLE}>Preset</span>
              <select
                style={DIALOG_INPUT_STYLE}
                data-testid="cs-format-cells-number-preset"
                value={state.numberFormatKey}
                onChange={(e) => {
                  const key = e.target.value as NumberFormatKey;
                  update('numberFormatKey', key);
                  update('numberPattern', NUMBER_FORMAT_PATTERNS[key]);
                }}
              >
                {NUMBER_FORMAT_OPTIONS.map((opt) => (
                  <option key={opt.value} value={opt.value}>
                    {opt.label}
                  </option>
                ))}
              </select>
            </label>
            <label style={DIALOG_FIELD_STYLE}>
              <span style={DIALOG_LABEL_STYLE}>Pattern</span>
              <input
                style={DIALOG_INPUT_STYLE}
                data-testid="cs-format-cells-number-pattern"
                value={state.numberPattern}
                onChange={(e) => update('numberPattern', e.target.value)}
              />
            </label>
          </>
        )}

        {tab === 'alignment' && (
          <>
            <label style={DIALOG_FIELD_STYLE}>
              <span style={DIALOG_LABEL_STYLE}>Horizontal</span>
              <select
                style={DIALOG_INPUT_STYLE}
                data-testid="cs-format-cells-align"
                value={state.align}
                onChange={(e) => update('align', e.target.value as HAlign)}
              >
                <option value="left">Left</option>
                <option value="center">Center</option>
                <option value="right">Right</option>
              </select>
            </label>
            <label style={DIALOG_FIELD_STYLE}>
              <span style={DIALOG_LABEL_STYLE}>Vertical</span>
              <select
                style={DIALOG_INPUT_STYLE}
                data-testid="cs-format-cells-valign"
                value={state.vAlign}
                onChange={(e) => update('vAlign', e.target.value as VAlign)}
              >
                <option value="top">Top</option>
                <option value="middle">Middle</option>
                <option value="bottom">Bottom</option>
              </select>
            </label>
            <label style={CHECK_STYLE} data-testid="cs-format-cells-wrap-label">
              <input
                type="checkbox"
                data-testid="cs-format-cells-wrap"
                checked={state.wrap}
                onChange={(e) => update('wrap', e.target.checked)}
              />
              <span>Wrap text</span>
            </label>
          </>
        )}

        {tab === 'font' && (
          <>
            <label style={DIALOG_FIELD_STYLE}>
              <span style={DIALOG_LABEL_STYLE}>Font</span>
              <select
                style={DIALOG_INPUT_STYLE}
                data-testid="cs-format-cells-font-family"
                value={state.fontFamily}
                onChange={(e) => update('fontFamily', e.target.value)}
              >
                {FONT_FAMILIES.map((family) => (
                  <option key={family} value={family}>
                    {family}
                  </option>
                ))}
              </select>
            </label>
            <label style={DIALOG_FIELD_STYLE}>
              <span style={DIALOG_LABEL_STYLE}>Size</span>
              <input
                style={DIALOG_INPUT_STYLE}
                data-testid="cs-format-cells-font-size"
                type="number"
                min={6}
                max={72}
                value={state.fontSize}
                onChange={(e) => update('fontSize', e.target.value)}
              />
            </label>
            <label style={CHECK_STYLE}>
              <input
                type="checkbox"
                data-testid="cs-format-cells-bold"
                checked={state.bold}
                onChange={(e) => update('bold', e.target.checked)}
              />
              <span>Bold</span>
            </label>
            <label style={CHECK_STYLE}>
              <input
                type="checkbox"
                data-testid="cs-format-cells-italic"
                checked={state.italic}
                onChange={(e) => update('italic', e.target.checked)}
              />
              <span>Italic</span>
            </label>
            <label style={CHECK_STYLE}>
              <input
                type="checkbox"
                data-testid="cs-format-cells-underline"
                checked={state.underline}
                onChange={(e) => update('underline', e.target.checked)}
              />
              <span>Underline</span>
            </label>
          </>
        )}

        {tab === 'border' && (
          <>
            <label style={DIALOG_FIELD_STYLE}>
              <span style={DIALOG_LABEL_STYLE}>Preset</span>
              <select
                style={DIALOG_INPUT_STYLE}
                data-testid="cs-format-cells-border-choice"
                value={state.borderChoice}
                onChange={(e) => update('borderChoice', e.target.value as BorderChoice)}
              >
                <option value="none">No border</option>
                <option value="all">All borders</option>
                <option value="outside">Outside border</option>
                <option value="top">Top border</option>
                <option value="bottom">Bottom border</option>
                <option value="left">Left border</option>
                <option value="right">Right border</option>
              </select>
            </label>
            <label style={DIALOG_FIELD_STYLE}>
              <span style={DIALOG_LABEL_STYLE}>Color</span>
              <input
                style={COLOR_INPUT_STYLE}
                data-testid="cs-format-cells-border-color"
                type="color"
                value={state.borderColor}
                onChange={(e) => update('borderColor', e.target.value)}
              />
            </label>
          </>
        )}

        {tab === 'fill' && (
          <label style={DIALOG_FIELD_STYLE}>
            <span style={DIALOG_LABEL_STYLE}>Fill color</span>
            <input
              style={COLOR_INPUT_STYLE}
              data-testid="cs-format-cells-fill-color"
              type="color"
              value={state.fillColor}
              onChange={(e) => update('fillColor', e.target.value)}
            />
          </label>
        )}
      </div>
    </Dialog>
  );
}
