import { useEffect, useState } from 'react';
import { Dialog } from './Dialog';
import { useUniverAPI } from '../use-univer';
import { useActiveCellState, type HAlign, type VAlign } from '../hooks/useActiveCellState';
import {
  DEFAULT_BORDER_COLOR,
  NUMBER_FORMAT_PATTERNS,
  setAlignment,
  setBold,
  setBorders,
  setFillColor,
  setFontFamily,
  setFontSize,
  setItalic,
  setNumberFormat,
  setUnderline,
  setVerticalAlignment,
  setWrap,
  type BorderChoice,
  type NumberFormatKey,
} from './home-tab-actions';

type Props = { onClose: () => void };

type TabId = 'number' | 'alignment' | 'font' | 'border' | 'fill';

type DialogState = {
  numberFormatKey: NumberFormatKey;
  numberPattern: string;
  align: Exclude<HAlign, 'unset'>;
  vAlign: Exclude<VAlign, 'unset'>;
  wrap: boolean;
  fontFamily: string;
  fontSize: string;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  borderChoice: BorderChoice;
  borderColor: string;
  fillColor: string;
};

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

function hasAnyBorder(style: unknown): boolean {
  if (!style || typeof style !== 'object' || !('bd' in style)) return false;
  const bd = (style as { bd?: Record<string, unknown> }).bd;
  return !!bd && Object.keys(bd).length > 0;
}

function buildInitialState(state: ReturnType<typeof useActiveCellState>, style: unknown): DialogState {
  const numberPattern = state.numberFormat;
  return {
    numberFormatKey: detectFormatKey(numberPattern),
    numberPattern,
    align: state.align === 'unset' ? 'left' : state.align,
    vAlign: state.vAlign === 'unset' ? 'bottom' : state.vAlign,
    wrap: state.isWrapped,
    fontFamily: state.fontFamily || 'Calibri',
    fontSize: String(state.fontSize || 11),
    bold: state.isBold,
    italic: state.isItalic,
    underline: state.isUnderline,
    borderChoice: hasAnyBorder(style) ? 'outside' : 'none',
    borderColor: DEFAULT_BORDER_COLOR,
    fillColor: state.fillColor || '#ffffff',
  };
}

export function FormatCellsDialog({ onClose }: Props) {
  const api = useUniverAPI();
  const active = useActiveCellState();
  const [tab, setTab] = useState<TabId>('number');
  const [dialogState, setDialogState] = useState<DialogState>(() => buildInitialState(active, null));

  useEffect(() => {
    if (!api || !active.ready) return;
    const wb = api.getActiveWorkbook();
    const sheet = wb?.getActiveSheet();
    const range = sheet?.getActiveRange();
    const cell = range ? sheet?.getRange(range.getRow(), range.getColumn()).getCellData() : null;
    const style =
      typeof cell?.s === 'string' ? (wb?.getWorkbook().getStyles().get(cell.s) ?? null) : (cell?.s ?? null);
    setDialogState(buildInitialState(active, style));
  }, [api, active]);

  const update = <K extends keyof DialogState>(key: K, value: DialogState[K]) =>
    setDialogState((prev) => ({ ...prev, [key]: value }));

  const apply = () => {
    if (!api) return;
    setNumberFormat(api, dialogState.numberPattern);
    setAlignment(api, dialogState.align);
    setVerticalAlignment(api, dialogState.vAlign);
    setWrap(api, dialogState.wrap);
    setFontFamily(api, dialogState.fontFamily);
    setFontSize(api, Number(dialogState.fontSize));
    setBold(api, dialogState.bold);
    setItalic(api, dialogState.italic);
    setUnderline(api, dialogState.underline);
    setBorders(api, dialogState.borderChoice, dialogState.borderColor);
    setFillColor(api, dialogState.fillColor);
    onClose();
  };

  return (
    <Dialog
      title="Format Cells"
      onClose={onClose}
      data-testid="format-cells-dialog"
      footer={
        <>
          <button type="button" className="btn-secondary" onClick={onClose}>
            Cancel
          </button>
          <button
            type="button"
            className="btn-primary"
            data-testid="format-cells-apply"
            onClick={apply}
          >
            Apply
          </button>
        </>
      }
    >
      <div className="format-cells">
        <div className="format-cells__tabs" role="tablist" aria-label="Format sections">
          {TABS.map((t) => (
            <button
              key={t.id}
              type="button"
              role="tab"
              className={`format-cells__tab${tab === t.id ? ' format-cells__tab--active' : ''}`}
              aria-selected={tab === t.id}
              data-testid={`format-cells-tab-${t.id}`}
              onClick={() => setTab(t.id)}
            >
              {t.label}
            </button>
          ))}
        </div>

        <div className="format-cells__panel" role="tabpanel">
          {tab === 'number' && (
            <section className="format-cells__section">
              <div className="field">
                <label className="field__label" htmlFor="format-cells-number-preset">
                  Preset
                </label>
                <select
                  id="format-cells-number-preset"
                  className="input"
                  data-testid="format-cells-number-preset"
                  value={dialogState.numberFormatKey}
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
              </div>
              <div className="field">
                <label className="field__label" htmlFor="format-cells-number-pattern">
                  Pattern
                </label>
                <input
                  id="format-cells-number-pattern"
                  className="input"
                  data-testid="format-cells-number-pattern"
                  value={dialogState.numberPattern}
                  onChange={(e) => update('numberPattern', e.target.value)}
                />
              </div>
            </section>
          )}

          {tab === 'alignment' && (
            <section className="format-cells__section">
              <div className="field">
                <label className="field__label" htmlFor="format-cells-align">
                  Horizontal
                </label>
                <select
                  id="format-cells-align"
                  className="input"
                  data-testid="format-cells-align"
                  value={dialogState.align}
                  onChange={(e) => update('align', e.target.value as DialogState['align'])}
                >
                  <option value="left">Left</option>
                  <option value="center">Center</option>
                  <option value="right">Right</option>
                </select>
              </div>
              <div className="field">
                <label className="field__label" htmlFor="format-cells-valign">
                  Vertical
                </label>
                <select
                  id="format-cells-valign"
                  className="input"
                  data-testid="format-cells-valign"
                  value={dialogState.vAlign}
                  onChange={(e) => update('vAlign', e.target.value as DialogState['vAlign'])}
                >
                  <option value="top">Top</option>
                  <option value="middle">Middle</option>
                  <option value="bottom">Bottom</option>
                </select>
              </div>
              <label className="format-cells__check" data-testid="format-cells-wrap-label">
                <input
                  type="checkbox"
                  data-testid="format-cells-wrap"
                  checked={dialogState.wrap}
                  onChange={(e) => update('wrap', e.target.checked)}
                />
                <span>Wrap text</span>
              </label>
            </section>
          )}

          {tab === 'font' && (
            <section className="format-cells__section">
              <div className="field">
                <label className="field__label" htmlFor="format-cells-font-family">
                  Font
                </label>
                <select
                  id="format-cells-font-family"
                  className="input"
                  data-testid="format-cells-font-family"
                  value={dialogState.fontFamily}
                  onChange={(e) => update('fontFamily', e.target.value)}
                >
                  {FONT_FAMILIES.map((family) => (
                    <option key={family} value={family}>
                      {family}
                    </option>
                  ))}
                </select>
              </div>
              <div className="field">
                <label className="field__label" htmlFor="format-cells-font-size">
                  Size
                </label>
                <input
                  id="format-cells-font-size"
                  className="input"
                  data-testid="format-cells-font-size"
                  type="number"
                  min={6}
                  max={72}
                  value={dialogState.fontSize}
                  onChange={(e) => update('fontSize', e.target.value)}
                />
              </div>
              <div className="format-cells__checks">
                <label className="format-cells__check">
                  <input
                    type="checkbox"
                    data-testid="format-cells-bold"
                    checked={dialogState.bold}
                    onChange={(e) => update('bold', e.target.checked)}
                  />
                  <span>Bold</span>
                </label>
                <label className="format-cells__check">
                  <input
                    type="checkbox"
                    data-testid="format-cells-italic"
                    checked={dialogState.italic}
                    onChange={(e) => update('italic', e.target.checked)}
                  />
                  <span>Italic</span>
                </label>
                <label className="format-cells__check">
                  <input
                    type="checkbox"
                    data-testid="format-cells-underline"
                    checked={dialogState.underline}
                    onChange={(e) => update('underline', e.target.checked)}
                  />
                  <span>Underline</span>
                </label>
              </div>
            </section>
          )}

          {tab === 'border' && (
            <section className="format-cells__section">
              <div className="field">
                <label className="field__label" htmlFor="format-cells-border-choice">
                  Preset
                </label>
                <select
                  id="format-cells-border-choice"
                  className="input"
                  data-testid="format-cells-border-choice"
                  value={dialogState.borderChoice}
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
              </div>
              <div className="field">
                <label className="field__label" htmlFor="format-cells-border-color">
                  Color
                </label>
                <input
                  id="format-cells-border-color"
                  className="format-cells__color"
                  data-testid="format-cells-border-color"
                  type="color"
                  value={dialogState.borderColor}
                  onChange={(e) => update('borderColor', e.target.value)}
                />
              </div>
            </section>
          )}

          {tab === 'fill' && (
            <section className="format-cells__section">
              <div className="field">
                <label className="field__label" htmlFor="format-cells-fill-color">
                  Fill color
                </label>
                <input
                  id="format-cells-fill-color"
                  className="format-cells__color"
                  data-testid="format-cells-fill-color"
                  type="color"
                  value={dialogState.fillColor}
                  onChange={(e) => update('fillColor', e.target.value)}
                />
              </div>
            </section>
          )}
        </div>
      </div>
    </Dialog>
  );
}
