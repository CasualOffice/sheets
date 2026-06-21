/**
 * Toolbar — the rich built-in formatting toolbar for `<CasualSheets chrome>`.
 *
 * Drives the editor purely through `CasualSheetsAPI.executeCommand` + the FUniver
 * facade (`api.univer`), with Material Symbols icons + design-system tokens. No
 * app context, no title/logo row (the host frames the editor with its own bar).
 * Commands grouped with dividers, Office-style — mirrors the app shell's
 * `apps/web/src/shell/Toolbar.tsx` + `RibbonControls.tsx` group/order/icons,
 * replicating `home-tab-actions.ts` logic inline against `api` / `api.univer`.
 *
 * Covered (parity with the app's Home tab):
 *   history (undo/redo) · clipboard (paste/cut/copy/paste-values) · format
 *   painter · font family/size + grow/shrink · bold/italic/underline/strike ·
 *   text + fill colour (ColorPicker) · borders (BordersPicker) · horizontal +
 *   vertical align · wrap · merge/unmerge · number-format dropdown +
 *   currency/percent/decimals · AutoSum (AutoSumPicker) · clear formatting.
 *
 * Reflects the active cell — toggles light up, dropdowns show the current font /
 * number format — via a command-execution subscription.
 *
 * Dialog-only controls (Format Cells, Insert Chart, PivotTable) have no SDK
 * dialog: when the host passes `onDialogRequest` they render and call it so the
 * host can render its own; without it, they are omitted entirely (no fake
 * dialog). Each group/control can also be hidden via the `features` flag map.
 */

import { useEffect, useState, type CSSProperties } from 'react';
import { ICommandService } from '@univerjs/core';
import type { CasualSheetsAPI } from '../sheets/api';
import { Icon } from './Icon';
import { ensureChromeFonts } from './fonts';
import { ColorPicker } from './ColorPicker';
import { BordersPicker } from './BordersPicker';
import { AutoSumPicker } from './AutoSumPicker';
import { useDialogs } from './dialog-context';
import type { ChromeExtensions, ToolbarExtension } from './extensions';

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
  nf: string; // number-format pattern ('' = General)
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
  nf: '',
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
    nf: typeof s?.n?.pattern === 'string' ? s.n.pattern : '',
  };
}

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
const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 72];

/**
 * Number-format presets, mirroring the app's `home-tab-actions`
 * `NUMBER_FORMAT_PATTERNS`. The dropdown applies a pattern via the facade
 * (`FRange.setNumberFormat`, a runtime numfmt-facade extension) so it lands as a
 * normal numfmt mutation — same path the app uses.
 */
const NUMBER_FORMAT_PATTERNS = {
  general: '',
  integer: '#,##0',
  number: '#,##0.00',
  currency: '"$"#,##0.00',
  accounting: '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)',
  percent: '0.00%',
  date: 'yyyy-mm-dd',
  time: 'hh:mm:ss',
  scientific: '0.00E+00',
  text: '@',
} as const;

type NumberFormatKey = keyof typeof NUMBER_FORMAT_PATTERNS;

const NUMBER_FORMAT_OPTIONS: { value: NumberFormatKey; label: string }[] = [
  { value: 'general', label: 'General' },
  { value: 'integer', label: 'Number (no decimals)' },
  { value: 'number', label: 'Number (2 decimals)' },
  { value: 'currency', label: 'Currency' },
  { value: 'accounting', label: 'Accounting' },
  { value: 'percent', label: 'Percent' },
  { value: 'date', label: 'Date (yyyy-mm-dd)' },
  { value: 'time', label: 'Time (hh:mm:ss)' },
  { value: 'scientific', label: 'Scientific' },
  { value: 'text', label: 'Text' },
];

function detectFormatKey(pattern: string): NumberFormatKey {
  for (const k of Object.keys(NUMBER_FORMAT_PATTERNS) as NumberFormatKey[]) {
    if (NUMBER_FORMAT_PATTERNS[k] === pattern) return k;
  }
  return 'general';
}

interface ToolbarAction {
  id: string;
  label: string;
  command: string;
  icon: string;
  params?: object;
  /** Feature key gating this action; defaults to the owning group's key. */
  feature?: string;
}

/**
 * Each entry is a feature-gated group. The `feature` key lets a host hide the
 * whole group (and per-action `feature` keys gate individual buttons). Order +
 * icons mirror the app shell's Home tab.
 *
 * HorizontalAlign enum (@univerjs/core): LEFT=1, CENTER=2, RIGHT=3.
 * VerticalAlign enum (@univerjs/core): TOP=1, MIDDLE=2, BOTTOM=3.
 * WrapStrategy enum (@univerjs/core): OVERFLOW=1, CLIP=2, WRAP=3.
 *
 * Clipboard command ids: `univer.command.{paste,cut,copy}` (matches
 * `home-tab-actions.ts`). Paste values-only uses the base paste command with a
 * predefined hook `value` (`special-paste-value`), same as the app's
 * `pasteSpecial`.
 */
const GROUPS: { feature: string; actions: ToolbarAction[] }[] = [
  {
    feature: 'history',
    actions: [
      { id: 'undo', label: 'Undo (Ctrl+Z)', command: 'univer.command.undo', icon: 'undo' },
      { id: 'redo', label: 'Redo (Ctrl+Y)', command: 'univer.command.redo', icon: 'redo' },
    ],
  },
  {
    feature: 'clipboard',
    actions: [
      {
        id: 'paste',
        label: 'Paste (Ctrl+V)',
        command: 'univer.command.paste',
        icon: 'content_paste',
      },
      { id: 'cut', label: 'Cut (Ctrl+X)', command: 'univer.command.cut', icon: 'content_cut' },
      { id: 'copy', label: 'Copy (Ctrl+C)', command: 'univer.command.copy', icon: 'content_copy' },
      {
        id: 'paste-values',
        label: 'Paste values only',
        command: 'univer.command.paste',
        icon: 'content_paste_go',
        params: { value: 'special-paste-value' },
      },
    ],
  },
  {
    feature: 'format-painter',
    actions: [
      {
        id: 'format-painter',
        label: 'Format Painter',
        command: 'sheet.command.set-once-format-painter',
        icon: 'format_paint',
      },
    ],
  },
  {
    feature: 'font-style',
    actions: [
      {
        id: 'bold',
        label: 'Bold (Ctrl+B)',
        command: 'sheet.command.set-range-bold',
        icon: 'format_bold',
      },
      {
        id: 'italic',
        label: 'Italic (Ctrl+I)',
        command: 'sheet.command.set-range-italic',
        icon: 'format_italic',
      },
      {
        id: 'underline',
        label: 'Underline (Ctrl+U)',
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
  },
  {
    feature: 'alignment',
    actions: [
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
      {
        id: 'wrap-text',
        label: 'Wrap text',
        command: 'sheet.command.set-text-wrap',
        icon: 'wrap_text',
        params: { value: 3 },
      },
    ],
  },
  {
    feature: 'alignment',
    actions: [
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
    ],
  },
  {
    feature: 'merge',
    actions: [
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
  },
  {
    feature: 'number',
    actions: [
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
  },
  {
    feature: 'clear-format',
    actions: [
      {
        id: 'clear-format',
        label: 'Clear formatting',
        command: 'sheet.command.clear-selection-format',
        icon: 'format_clear',
      },
    ],
  },
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
  flexWrap: 'wrap',
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
  /**
   * Per-control / per-group feature flags. A key set to `false` hides that
   * control or group; omitted keys default to enabled. Keys:
   *   history · clipboard · format-painter · font · font-style · color ·
   *   borders · alignment · merge · number · autosum · clear-format ·
   *   format-cells · insert-chart · pivot-table
   */
  features?: Record<string, boolean>;
  /** Host chrome extensions — only `extensions.toolbar` is read here. */
  extensions?: ChromeExtensions;
}

/** A feature is on unless explicitly set to `false`. */
function enabled(features: Record<string, boolean> | undefined, key: string): boolean {
  return features?.[key] !== false;
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

export function Toolbar({ api, features, extensions }: ToolbarProps) {
  const dialogs = useDialogs();
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

  /**
   * Apply a font size via the dedicated command (matches the app's
   * `set-range-fontsize`). Used by the size dropdown + grow/shrink.
   */
  const applyFontSize = (size: number) => {
    if (!Number.isFinite(size) || size <= 0) return;
    dispatch('sheet.command.set-range-fontsize', { value: size });
  };

  /**
   * Grow / shrink the active cell's font size by `delta`, mirroring the app's
   * `adjustFontSize` (read current off the active cell, fall back to 11, clamp
   * to [6, 72]). Reads the size we already track in `active`.
   */
  const adjustFontSize = (delta: number) => {
    const current = active.fs && active.fs > 0 ? active.fs : 11;
    const next = Math.max(6, Math.min(72, current + delta));
    if (next === current) return;
    applyFontSize(next);
  };

  /**
   * Apply a number-format pattern via the FUniver facade — `FRange.setNumberFormat`
   * is a runtime numfmt-facade extension (added via `FUniver.extend()`), so a
   * runtime cast is the cleanest way to reach it. Same approach as the app's
   * `home-tab-actions.setNumberFormat`.
   */
  const applyNumberFormat = (pattern: string) => {
    const sheet = api?.univer.getActiveWorkbook()?.getActiveSheet();
    const range = (sheet as unknown as { getActiveRange?: () => unknown })?.getActiveRange?.() as
      | { setNumberFormat?: (p: string) => unknown }
      | undefined;
    range?.setNumberFormat?.(pattern);
  };

  // Reflect current font in the dropdowns; surface a non-listed value too.
  const familyValue = active.ff || 'Calibri';
  const familyOptions =
    active.ff && !FONT_FAMILIES.includes(active.ff) ? [active.ff, ...FONT_FAMILIES] : FONT_FAMILIES;
  const sizeValue = active.fs || 11;
  const sizeOptions =
    active.fs && !FONT_SIZES.includes(active.fs) ? [active.fs, ...FONT_SIZES] : FONT_SIZES;
  const numberFormatValue = detectFormatKey(active.nf);

  const renderActionButton = (a: ToolbarAction) => {
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
        data-testid={`cs-${a.id}`}
        data-active={on ? 'true' : undefined}
        disabled={!api}
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
          if (!on) e.currentTarget.style.background = 'var(--cs-chrome-hover, rgba(0,0,0,0.06))';
        }}
        onMouseLeave={(e) => {
          e.currentTarget.style.background = baseBg;
        }}
      >
        <Icon name={a.icon} size={20} />
      </button>
    );
  };

  // A plain (non-command) icon button — used for grow/shrink + dialog requests.
  const renderIconButton = (
    id: string,
    label: string,
    icon: string,
    onPress: () => void,
    testid: string,
  ) => (
    <button
      key={id}
      type="button"
      title={label}
      aria-label={label}
      data-action={id}
      data-testid={testid}
      disabled={!api}
      style={BTN_STYLE}
      onMouseDown={(e) => {
        e.preventDefault();
        onPress();
      }}
      onMouseEnter={(e) => {
        e.currentTarget.style.background = 'var(--cs-chrome-hover, rgba(0,0,0,0.06))';
      }}
      onMouseLeave={(e) => {
        e.currentTarget.style.background = 'transparent';
      }}
    >
      <Icon name={icon} size={20} />
    </button>
  );

  // Render the static command groups, skipping any whose feature is off and any
  // group left empty after per-action filtering. A leading divider separates a
  // group from whatever rendered before it.
  let firstGroupRendered = false;
  const groupNodes = GROUPS.map((group, gi) => {
    if (!enabled(features, group.feature)) return null;
    const visible = group.actions.filter((a) => enabled(features, a.feature ?? group.feature));
    if (visible.length === 0) return null;
    const showDivider = firstGroupRendered;
    firstGroupRendered = true;
    return (
      <span key={gi} style={{ display: 'inline-flex', alignItems: 'center' }}>
        {showDivider && <span style={DIVIDER_STYLE} aria-hidden />}
        {visible.map(renderActionButton)}
      </span>
    );
  });

  const showFont = enabled(features, 'font');
  const showColor = enabled(features, 'color');
  const showBorders = enabled(features, 'borders');
  const showAutoSum = enabled(features, 'autosum');
  const showNumberDropdown = enabled(features, 'number');

  // Dialog-backed controls: render when the chrome can open them — Format Cells
  // is a built-in (always available); Insert Chart / PivotTable open only when a
  // host registered/handles them.
  const showFormatCells = enabled(features, 'format-cells') && dialogs.canOpen('format-cells');
  const showInsertChart = enabled(features, 'insert-chart') && dialogs.canOpen('insert-chart');
  const showPivotTable = enabled(features, 'pivot-table') && dialogs.canOpen('insert-pivot');
  const showDialogGroup = showFormatCells || showInsertChart || showPivotTable;

  // Host toolbar extensions — appended after the built-in groups.
  const toolbarExt = (extensions?.toolbar ?? []).filter(
    (e) => !e.isVisible || (api ? e.isVisible(api) : false),
  );
  const runExt = (e: ToolbarExtension) => {
    if (!api) return;
    if (e.command) void api.executeCommand(e.command, e.commandParams);
    else e.onClick?.(api);
  };

  return (
    <div style={BAR_STYLE} data-testid="casual-sheets-toolbar" role="toolbar" aria-label="Editor">
      {showFont && (
        <>
          <select
            aria-label="Font family"
            data-testid="cs-font-family"
            style={{ ...SELECT_STYLE, width: 130 }}
            value={familyValue}
            disabled={!api}
            onChange={(e) =>
              dispatch('sheet.command.set-range-font-family', { value: e.target.value })
            }
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
            disabled={!api}
            onChange={(e) => applyFontSize(Number(e.target.value))}
          >
            {sizeOptions.map((s) => (
              <option key={s} value={s}>
                {s}
              </option>
            ))}
          </select>
          {renderIconButton(
            'font-size-up',
            'Increase font size',
            'text_increase',
            () => adjustFontSize(+1),
            'cs-font-size-up',
          )}
          {renderIconButton(
            'font-size-down',
            'Decrease font size',
            'text_decrease',
            () => adjustFontSize(-1),
            'cs-font-size-down',
          )}
          <span style={DIVIDER_STYLE} aria-hidden />
        </>
      )}

      {groupNodes}

      {(showColor || showBorders) && <span style={DIVIDER_STYLE} aria-hidden />}
      {showColor && <ColorPicker api={api} />}
      {showBorders && <BordersPicker api={api} />}

      {showNumberDropdown && (
        <>
          <span style={DIVIDER_STYLE} aria-hidden />
          <select
            aria-label="Number format"
            data-testid="cs-number-format"
            style={{ ...SELECT_STYLE, width: 124 }}
            value={numberFormatValue}
            disabled={!api}
            onChange={(e) =>
              applyNumberFormat(NUMBER_FORMAT_PATTERNS[e.target.value as NumberFormatKey])
            }
          >
            {NUMBER_FORMAT_OPTIONS.map((o) => (
              <option key={o.value} value={o.value}>
                {o.label}
              </option>
            ))}
          </select>
        </>
      )}

      {showAutoSum && (
        <>
          <span style={DIVIDER_STYLE} aria-hidden />
          <AutoSumPicker api={api} />
        </>
      )}

      {showDialogGroup && (
        <>
          <span style={DIVIDER_STYLE} aria-hidden />
          {showFormatCells &&
            renderIconButton(
              'format-cells',
              'Format Cells…',
              'format_shapes',
              () => dialogs.openDialog('format-cells'),
              'cs-format-cells',
            )}
          {showInsertChart &&
            renderIconButton(
              'insert-chart',
              'Insert chart',
              'bar_chart',
              () => dialogs.openDialog('insert-chart'),
              'cs-insert-chart',
            )}
          {showPivotTable &&
            renderIconButton(
              'pivot-table',
              'Insert PivotTable',
              'pivot_table_chart',
              () => dialogs.openDialog('insert-pivot'),
              'cs-pivot-table',
            )}
        </>
      )}

      {toolbarExt.length > 0 && (
        <>
          <span style={DIVIDER_STYLE} aria-hidden />
          {toolbarExt.map((e) =>
            renderIconButton(e.id, e.label, e.icon, () => runExt(e), `cs-ext-${e.id}`),
          )}
        </>
      )}
    </div>
  );
}
