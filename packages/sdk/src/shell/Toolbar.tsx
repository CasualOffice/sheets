import type { CSSProperties, ReactNode } from 'react';
import { IconButton, Select, type SelectOption } from '@schnsrw/design-system';

export interface ToolbarFormatState {
  fontFamily?: string | null;
  fontSize?: number | null;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  color?: string | null;
  fill?: string | null;
  align?: 'left' | 'center' | 'right';
  numfmt?: string;
}

export interface ToolbarCallbacks {
  onUndo?: () => void;
  onRedo?: () => void;
  onPaintFormat?: () => void;

  onSetFontFamily?: (family: string) => void;
  onSetFontSize?: (size: number) => void;

  onToggleBold?: () => void;
  onToggleItalic?: () => void;
  onToggleUnderline?: () => void;
  onToggleStrikethrough?: () => void;

  onSetTextColor?: () => void;
  onSetFillColor?: () => void;
  onSetBorders?: () => void;

  onAlignLeft?: () => void;
  onAlignCenter?: () => void;
  onAlignRight?: () => void;
  onToggleWrap?: () => void;
  onMergeCells?: () => void;

  onSetNumberFormat?: (format: string) => void;
  onFormatCurrency?: () => void;
  onFormatPercent?: () => void;
  onAutoSum?: () => void;

  onInsertChart?: () => void;
  onInsertPivot?: () => void;
  onFormatTable?: () => void;
  onToggleFilter?: () => void;
  onInsertComment?: () => void;
}

export interface ToolbarProps extends ToolbarCallbacks {
  fmt?: ToolbarFormatState;
  /** Override the suggested font family list. */
  fontFamilies?: string[];
  /** Override the suggested font size list. */
  fontSizes?: number[];
  /** Override the suggested number-format list. */
  numberFormats?: string[];
  style?: CSSProperties;
}

const DEFAULT_FONT_FAMILIES = [
  'Calibri',
  'Arial',
  'Inter',
  'Times New Roman',
  'Georgia',
  'JetBrains Mono',
];
const DEFAULT_FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48];
const DEFAULT_NUMBER_FORMATS = [
  'General',
  'Number',
  'Currency',
  'Accounting',
  'Percent',
  'Date',
  'Scientific',
  'Text',
];

export function Toolbar({
  fmt = {},
  fontFamilies = DEFAULT_FONT_FAMILIES,
  fontSizes = DEFAULT_FONT_SIZES,
  numberFormats = DEFAULT_NUMBER_FORMATS,
  style,
  ...cb
}: ToolbarProps) {
  const fontFamilyOptions: SelectOption[] = fontFamilies.map((f) => ({ value: f, label: f }));
  const fontSizeOptions: SelectOption[] = fontSizes.map((n) => ({
    value: String(n),
    label: String(n),
  }));
  const numfmtOptions: SelectOption[] = numberFormats.map((n) => ({ value: n, label: n }));

  return (
    <div
      style={{
        height: 50,
        flex: '0 0 50px',
        display: 'flex',
        alignItems: 'center',
        background: 'var(--color-surface)',
        borderBottom: '1px solid var(--color-divider)',
        padding: '0 14px',
        gap: 0,
        overflowX: 'auto',
        ...style,
      }}
    >
      <Cluster>
        <IconButton icon="undo" label="Undo (Ctrl+Z)" onClick={cb.onUndo} />
        <IconButton icon="redo" label="Redo (Ctrl+Y)" onClick={cb.onRedo} />
        <IconButton icon="format_paint" label="Format painter" onClick={cb.onPaintFormat} />
      </Cluster>
      <Sep />

      <div style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
        <Select
          options={fontFamilyOptions}
          value={fmt.fontFamily ?? 'Calibri'}
          width={124}
          onChange={(e) => cb.onSetFontFamily?.(e.currentTarget.value)}
        />
        <Select
          options={fontSizeOptions}
          value={String(fmt.fontSize ?? 11)}
          width={58}
          onChange={(e) => cb.onSetFontSize?.(Number(e.currentTarget.value))}
        />
      </div>
      <Sep />

      <Cluster>
        <IconButton
          icon="format_bold"
          label="Bold (Ctrl+B)"
          pressed={!!fmt.bold}
          onClick={cb.onToggleBold}
        />
        <IconButton
          icon="format_italic"
          label="Italic (Ctrl+I)"
          pressed={!!fmt.italic}
          onClick={cb.onToggleItalic}
        />
        <IconButton
          icon="format_underlined"
          label="Underline (Ctrl+U)"
          pressed={!!fmt.underline}
          onClick={cb.onToggleUnderline}
        />
        <IconButton
          icon="format_strikethrough"
          label="Strikethrough"
          pressed={!!fmt.strikethrough}
          onClick={cb.onToggleStrikethrough}
        />
      </Cluster>
      <Sep />

      <Cluster>
        <IconButton
          icon="format_color_text"
          label="Font color"
          pressed={!!fmt.color}
          onClick={cb.onSetTextColor}
        />
        <IconButton
          icon="format_color_fill"
          label="Fill color"
          pressed={!!fmt.fill}
          onClick={cb.onSetFillColor}
        />
        <IconButton icon="border_all" label="Borders" onClick={cb.onSetBorders} />
      </Cluster>
      <Sep />

      <Cluster>
        <IconButton
          icon="format_align_left"
          label="Align left"
          pressed={(fmt.align ?? 'left') === 'left'}
          onClick={cb.onAlignLeft}
        />
        <IconButton
          icon="format_align_center"
          label="Center"
          pressed={fmt.align === 'center'}
          onClick={cb.onAlignCenter}
        />
        <IconButton
          icon="format_align_right"
          label="Align right"
          pressed={fmt.align === 'right'}
          onClick={cb.onAlignRight}
        />
        <IconButton icon="wrap_text" label="Wrap text" onClick={cb.onToggleWrap} />
        <IconButton icon="cell_merge" label="Merge and center" onClick={cb.onMergeCells} />
      </Cluster>
      <Sep />

      <div style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
        <Select
          options={numfmtOptions}
          value={fmt.numfmt ?? 'General'}
          width={106}
          onChange={(e) => cb.onSetNumberFormat?.(e.currentTarget.value)}
        />
        <Cluster>
          <IconButton icon="attach_money" label="Currency" onClick={cb.onFormatCurrency} />
          <IconButton icon="percent" label="Percent" onClick={cb.onFormatPercent} />
          <IconButton icon="functions" label="Auto sum" onClick={cb.onAutoSum} />
        </Cluster>
      </div>
      <Sep />

      <Cluster>
        <IconButton icon="bar_chart" label="Insert chart" onClick={cb.onInsertChart} />
        <IconButton icon="pivot_table_chart" label="Insert PivotTable" onClick={cb.onInsertPivot} />
        <IconButton icon="table_rows" label="Format as table" onClick={cb.onFormatTable} />
        <IconButton icon="filter_alt" label="Toggle filter" onClick={cb.onToggleFilter} />
        <IconButton icon="add_comment" label="Insert comment" onClick={cb.onInsertComment} />
      </Cluster>
    </div>
  );
}

function Cluster({ children }: { children: ReactNode }) {
  return (
    <div
      style={{
        display: 'inline-flex',
        alignItems: 'center',
        gap: 1,
        padding: 2,
        background: 'var(--color-surface-alt)',
        borderRadius: 'var(--radius-md)',
      }}
    >
      {children}
    </div>
  );
}

function Sep() {
  return (
    <span
      style={{
        width: 1,
        height: 22,
        background: 'var(--color-divider)',
        margin: '0 5px',
        flex: '0 0 auto',
      }}
    />
  );
}
