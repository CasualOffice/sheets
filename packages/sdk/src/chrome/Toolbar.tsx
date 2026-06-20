/**
 * Toolbar — the rich built-in formatting toolbar for `<CasualSheets chrome>`.
 *
 * Drives the editor purely through `CasualSheetsAPI.executeCommand`, with
 * Material Symbols icons + design-system tokens. No app context, no title/logo
 * row (the host frames the editor with its own bar). Commands grouped with
 * dividers, Office-style.
 *
 * Covered: font family/size · undo/redo · bold/italic/underline/strikethrough ·
 * horizontal align · merge/unmerge · number formats (currency/percent/decimals).
 * Reflects the active cell — toggles light up, dropdowns show the current font —
 * via a command-execution subscription. Follow-up: text/fill colour pickers
 * (need a swatch popover).
 */

import { useEffect, useState, type CSSProperties } from 'react';
import { ICommandService } from '@univerjs/core';
import type { CasualSheetsAPI } from '../sheets/api';
import { Icon } from './Icon';
import { ensureChromeFonts } from './fonts';
import { ColorPicker } from './ColorPicker';
import { BordersPicker } from './BordersPicker';
import { AutoSumPicker } from './AutoSumPicker';

interface ActiveStyle {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  ht: number; // HorizontalAlign of the active cell (0 = unset)
  vt: number; // VerticalAlign of the active cell (0 = unset)
  tb: number; // WrapStrategy of the active cell (0 = unset, 3 = wrap)
  ff: string;
  fs: number;
}

const NO_STYLE: ActiveStyle = {
  bold: false,
  italic: false,
  underline: false,
  strike: false,
  ht: 0,
  vt: 0,
  tb: 0,
  ff: '',
  fs: 0,
};

function readActiveStyle(api: CasualSheetsAPI): ActiveStyle {
  const sel = api.getSelection();
  const sheet = api.univer.getActiveWorkbook()?.getActiveSheet();
  if (!sel || !sheet) return NO_STYLE;
  const range = sheet.getRange(sel.range.startRow, sel.range.startColumn);
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const s = (range.getCellStyleData?.() ?? null) as any;
  return {
    bold: s?.bl === 1,
    italic: s?.it === 1,
    underline: s?.ul?.s === 1,
    strike: s?.st?.s === 1,
    ht: typeof s?.ht === 'number' ? s.ht : 0,
    vt: typeof s?.vt === 'number' ? s.vt : 0,
    tb: typeof s?.tb === 'number' ? s.tb : 0,
    ff: typeof s?.ff === 'string' ? s.ff : '',
    fs: typeof s?.fs === 'number' ? s.fs : 0,
  };
}

const FONT_FAMILIES = ['Arial', 'Calibri', 'Times New Roman', 'Courier New', 'Georgia', 'Verdana'];
const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72];

interface ToolbarAction {
  id: string;
  label: string;
  command: string;
  icon: string;
  params?: object;
}

// HorizontalAlign enum (@univerjs/core): LEFT=1, CENTER=2, RIGHT=3.
// VerticalAlign enum (@univerjs/core): TOP=1, MIDDLE=2, BOTTOM=3.
// WrapStrategy enum (@univerjs/core): OVERFLOW=1, CLIP=2, WRAP=3.
const GROUPS: ToolbarAction[][] = [
  [
    { id: 'undo', label: 'Undo', command: 'univer.command.undo', icon: 'undo' },
    { id: 'redo', label: 'Redo', command: 'univer.command.redo', icon: 'redo' },
  ],
  [
    { id: 'bold', label: 'Bold', command: 'sheet.command.set-range-bold', icon: 'format_bold' },
    {
      id: 'italic',
      label: 'Italic',
      command: 'sheet.command.set-range-italic',
      icon: 'format_italic',
    },
    {
      id: 'underline',
      label: 'Underline',
      command: 'sheet.command.set-range-underline',
      icon: 'format_underlined',
    },
    {
      id: 'strikethrough',
      label: 'Strikethrough',
      command: 'sheet.command.set-range-stroke',
      icon: 'format_strikethrough',
    },
  ],
  [
    {
      id: 'align-left',
      label: 'Align left',
      command: 'sheet.command.set-horizontal-text-align',
      icon: 'format_align_left',
      params: { value: 1 },
    },
    {
      id: 'align-center',
      label: 'Align center',
      command: 'sheet.command.set-horizontal-text-align',
      icon: 'format_align_center',
      params: { value: 2 },
    },
    {
      id: 'align-right',
      label: 'Align right',
      command: 'sheet.command.set-horizontal-text-align',
      icon: 'format_align_right',
      params: { value: 3 },
    },
  ],
  [
    {
      id: 'align-top',
      label: 'Align top',
      command: 'sheet.command.set-vertical-text-align',
      icon: 'vertical_align_top',
      params: { value: 1 },
    },
    {
      id: 'align-middle',
      label: 'Align middle',
      command: 'sheet.command.set-vertical-text-align',
      icon: 'vertical_align_center',
      params: { value: 2 },
    },
    {
      id: 'align-bottom',
      label: 'Align bottom',
      command: 'sheet.command.set-vertical-text-align',
      icon: 'vertical_align_bottom',
      params: { value: 3 },
    },
    {
      id: 'wrap-text',
      label: 'Wrap text',
      command: 'sheet.command.set-text-wrap',
      icon: 'wrap_text',
      params: { value: 3 },
    },
  ],
  [
    {
      id: 'merge',
      label: 'Merge cells',
      command: 'sheet.command.add-worksheet-merge-all',
      icon: 'cell_merge',
    },
    {
      id: 'unmerge',
      label: 'Unmerge cells',
      command: 'sheet.command.remove-worksheet-merge',
      icon: 'splitscreen_vertical_add',
    },
  ],
  [
    {
      id: 'currency',
      label: 'Currency format',
      command: 'sheet.command.numfmt.set.currency',
      icon: 'attach_money',
    },
    {
      id: 'percent',
      label: 'Percent format',
      command: 'sheet.command.numfmt.set.percent',
      icon: 'percent',
    },
    {
      id: 'decimal-increase',
      label: 'Increase decimals',
      command: 'sheet.command.numfmt.add.decimal.command',
      icon: 'decimal_increase',
    },
    {
      id: 'decimal-decrease',
      label: 'Decrease decimals',
      command: 'sheet.command.numfmt.subtract.decimal.command',
      icon: 'decimal_decrease',
    },
  ],
  [
    {
      id: 'clear-format',
      label: 'Clear formatting',
      command: 'sheet.command.clear-selection-format',
      icon: 'format_clear',
    },
  ],
];

const BAR_STYLE: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 2,
  padding: '4px 8px',
  borderBottom: '1px solid var(--cs-chrome-border, #e6e9ee)',
  background: 'var(--cs-chrome-bg, #eef1f5)',
  flex: '0 0 auto',
  userSelect: 'none',
};

const BTN_STYLE: CSSProperties = {
  width: 30,
  height: 30,
  display: 'inline-flex',
  alignItems: 'center',
  justifyContent: 'center',
  border: 'none',
  borderRadius: 6,
  background: 'transparent',
  cursor: 'pointer',
  color: 'var(--cs-chrome-fg, #201f1e)',
  padding: 0,
};

const DIVIDER_STYLE: CSSProperties = {
  width: 1,
  height: 18,
  margin: '0 4px',
  background: 'var(--cs-chrome-border, #cdd3db)',
  flex: '0 0 auto',
};

const SELECT_STYLE: CSSProperties = {
  height: 26,
  border: '1px solid var(--cs-chrome-border, #cdd3db)',
  borderRadius: 6,
  background: 'var(--cs-chrome-input-bg, #fff)',
  color: 'var(--cs-chrome-fg, #201f1e)',
  font: 'inherit',
  fontSize: 13,
  padding: '0 4px',
  cursor: 'pointer',
};

export interface ToolbarProps {
  /** Live API, or `null` until the editor is ready. */
  api: CasualSheetsAPI | null;
}

// Which actions reflect an active state, keyed off the active cell's style.
function isActive(id: string, s: ActiveStyle): boolean {
  switch (id) {
    case 'bold':
      return s.bold;
    case 'italic':
      return s.italic;
    case 'underline':
      return s.underline;
    case 'strikethrough':
      return s.strike;
    case 'align-left':
      return s.ht === 1;
    case 'align-center':
      return s.ht === 2;
    case 'align-right':
      return s.ht === 3;
    case 'align-top':
      return s.vt === 1;
    case 'align-middle':
      return s.vt === 2;
    case 'align-bottom':
      return s.vt === 3;
    case 'wrap-text':
      return s.tb === 3;
    default:
      return false;
  }
}

export function Toolbar({ api }: ToolbarProps) {
  const [active, setActive] = useState<ActiveStyle>(NO_STYLE);

  useEffect(() => {
    ensureChromeFonts();
  }, []);

  // Reflect the active cell: subscribe to command activity (covers selection
  // moves + style mutations) and re-read the style.
  useEffect(() => {
    if (!api) return;
    const refresh = () => setActive(readActiveStyle(api));
    refresh();
    const injector = (api.univer as unknown as { _injector?: { get(t: unknown): unknown } })
      ._injector;
    const cmd = injector?.get(ICommandService) as
      | { onCommandExecuted: (cb: () => void) => { dispose: () => void } }
      | undefined;
    const sub = cmd?.onCommandExecuted(() => refresh());
    return () => sub?.dispose();
  }, [api]);

  const dispatch = (command: string, params?: object) => void api?.executeCommand(command, params);

  // Reflect current font in the dropdowns; surface a non-listed value too.
  const familyValue = active.ff || 'Arial';
  const familyOptions =
    active.ff && !FONT_FAMILIES.includes(active.ff) ? [active.ff, ...FONT_FAMILIES] : FONT_FAMILIES;
  const sizeValue = active.fs || 11;
  const sizeOptions =
    active.fs && !FONT_SIZES.includes(active.fs) ? [active.fs, ...FONT_SIZES] : FONT_SIZES;

  return (
    <div style={BAR_STYLE} data-testid="casual-sheets-toolbar" role="toolbar" aria-label="Editor">
      <select
        aria-label="Font family"
        data-testid="cs-font-family"
        style={{ ...SELECT_STYLE, width: 116 }}
        value={familyValue}
        onChange={(e) => dispatch('sheet.command.set-range-font-family', { value: e.target.value })}
      >
        {familyOptions.map((f) => (
          <option key={f} value={f}>
            {f}
          </option>
        ))}
      </select>
      <select
        aria-label="Font size"
        data-testid="cs-font-size"
        style={{ ...SELECT_STYLE, width: 56, marginLeft: 4 }}
        value={sizeValue}
        onChange={(e) =>
          dispatch('sheet.command.set-range-fontsize', { value: Number(e.target.value) })
        }
      >
        {sizeOptions.map((s) => (
          <option key={s} value={s}>
            {s}
          </option>
        ))}
      </select>
      <span style={DIVIDER_STYLE} aria-hidden />
      {GROUPS.map((group, gi) => (
        <span key={gi} style={{ display: 'inline-flex', alignItems: 'center' }}>
          {gi > 0 && <span style={DIVIDER_STYLE} aria-hidden />}
          {group.map((a) => {
            const on = isActive(a.id, active);
            const baseBg = on ? 'var(--cs-chrome-active, #e6f3f7)' : 'transparent';
            return (
              <button
                key={a.id}
                type="button"
                title={a.label}
                aria-label={a.label}
                aria-pressed={on}
                data-action={a.id}
                data-active={on ? 'true' : undefined}
                style={{
                  ...BTN_STYLE,
                  background: baseBg,
                  color: on ? 'var(--cs-chrome-active-fg, #0e7490)' : BTN_STYLE.color,
                }}
                // mousedown (not click) so the grid's selection isn't lost first.
                onMouseDown={(e) => {
                  e.preventDefault();
                  dispatch(a.command, a.params);
                }}
                onMouseEnter={(e) => {
                  if (!on)
                    e.currentTarget.style.background = 'var(--cs-chrome-hover, rgba(0,0,0,0.06))';
                }}
                onMouseLeave={(e) => {
                  e.currentTarget.style.background = baseBg;
                }}
              >
                <Icon name={a.icon} size={20} />
              </button>
            );
          })}
        </span>
      ))}
      <span style={DIVIDER_STYLE} aria-hidden />
      <ColorPicker api={api} />
      <BordersPicker api={api} />
      <span style={DIVIDER_STYLE} aria-hidden />
      <AutoSumPicker api={api} />
    </div>
  );
}
