import { useUniverAPI } from '../use-univer';
import { useUI } from '../use-ui';
import { useActiveCellState, type HAlign, type VAlign } from '../hooks/useActiveCellState';
import {
  NUMBER_FORMATS,
  copy,
  cut,
  paste,
  redo,
  setAlignment,
  setBorders,
  setFillColor,
  setFontColor,
  setFontFamily,
  setFontSize,
  setNumberFormat,
  setNumberFormatByKey,
  setVerticalAlignment,
  startFormatPainter,
  toggleBold,
  toggleItalic,
  toggleMerge,
  toggleStrikethrough,
  toggleUnderline,
  toggleWrap,
  undo,
  type BorderChoice,
  type NumberFormatKey,
} from './home-tab-actions';
import {
  TABLE_THEMES,
  applyAutoFunction,
  formatAsTable,
  insertComment,
  insertHyperlink,
  sortRange,
  toggleFilter,
  type TableThemeId,
} from './tab-actions';
import {
  BordersControl,
  RibbonGroup,
  ToolbarButton,
  ToolbarColorButton,
  ToolbarDropdown,
  ToolbarSelect,
} from './RibbonControls';
import { DEFAULT_BORDER_COLOR } from './home-tab-actions';

/**
 * Fixed single-row toolbar — every group lays out its controls in one
 * horizontal row, so all icons share a baseline. No 2-row groups → no
 * visual asymmetry. Less-common features live in the top MenuBar.
 */

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
const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48];

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
  const patterns: Record<NumberFormatKey, string> = {
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
  };
  for (const k of Object.keys(patterns) as NumberFormatKey[]) if (patterns[k] === pattern) return k;
  return 'general';
}

export function Toolbar() {
  const api = useUniverAPI();
  const ui = useUI();
  const state = useActiveCellState();
  const ready = Boolean(api) && state.ready;

  return (
    <nav className="toolbar" data-testid="toolbar" aria-label="Toolbar">
      <div className="toolbar__inner">
        <RibbonGroup label="History">
          <ToolbarButton id="undo" label="Undo (Ctrl+Z)" icon="undo" disabled={!ready} onClick={() => api && undo(api)} />
          <ToolbarButton id="redo" label="Redo (Ctrl+Y)" icon="redo" disabled={!ready} onClick={() => api && redo(api)} />
        </RibbonGroup>

        <RibbonGroup label="Clipboard">
          <ToolbarButton id="paste" label="Paste (Ctrl+V)" icon="content_paste" disabled={!ready} onClick={() => api && paste(api)} />
          <ToolbarButton id="cut" label="Cut (Ctrl+X)" icon="content_cut" disabled={!ready} onClick={() => api && cut(api)} />
          <ToolbarButton id="copy" label="Copy (Ctrl+C)" icon="content_copy" disabled={!ready} onClick={() => api && copy(api)} />
          <ToolbarButton
            id="format-painter"
            label={state.isFormatPainterActive ? 'Format Painter (armed)' : 'Format Painter'}
            icon="format_paint"
            pressed={state.isFormatPainterActive}
            disabled={!ready}
            onClick={() => api && startFormatPainter(api)}
          />
        </RibbonGroup>

        <RibbonGroup label="Font">
          <ToolbarSelect
            id="font-family"
            label="Font family"
            value={state.fontFamily || 'Calibri'}
            options={FONT_FAMILIES.map((f) => ({ value: f, label: f }))}
            width={130}
            disabled={!ready}
            onChange={(v) => api && setFontFamily(api, v)}
          />
          <ToolbarSelect
            id="font-size"
            label="Font size"
            value={String(state.fontSize || 11)}
            options={FONT_SIZES.map((s) => ({ value: String(s), label: String(s) }))}
            width={56}
            disabled={!ready}
            onChange={(v) => api && setFontSize(api, Number(v))}
          />
          <ToolbarButton id="bold" label="Bold (Ctrl+B)" icon="format_bold" pressed={state.isBold} disabled={!ready} onClick={() => api && toggleBold(api, state.isBold)} />
          <ToolbarButton id="italic" label="Italic (Ctrl+I)" icon="format_italic" pressed={state.isItalic} disabled={!ready} onClick={() => api && toggleItalic(api, state.isItalic)} />
          <ToolbarButton id="underline" label="Underline (Ctrl+U)" icon="format_underlined" pressed={state.isUnderline} disabled={!ready} onClick={() => api && toggleUnderline(api, state.isUnderline)} />
          <ToolbarButton id="strikethrough" label="Strikethrough" icon="format_strikethrough" pressed={state.isStrike} disabled={!ready} onClick={() => api && toggleStrikethrough(api, state.isStrike)} />
          <ToolbarColorButton
            id="font-color"
            label="Font color"
            icon="format_color_text"
            value={state.fontColor}
            defaultColor="#000000"
            disabled={!ready}
            onChange={(c) => api && setFontColor(api, c)}
          />
          <ToolbarColorButton
            id="fill-color"
            label="Fill / highlight color"
            icon="format_color_fill"
            value={state.fillColor}
            defaultColor="#ffeb3b"
            disabled={!ready}
            onChange={(c) => api && setFillColor(api, c)}
          />
          <BordersControl
            label="Borders"
            icon="border_all"
            disabled={!ready}
            defaultColor={DEFAULT_BORDER_COLOR}
            items={[
              { id: 'all', label: 'All borders', icon: 'border_all' },
              { id: 'outside', label: 'Outside borders', icon: 'border_outer' },
              { id: 'top', label: 'Top border', icon: 'border_top' },
              { id: 'bottom', label: 'Bottom border', icon: 'border_bottom' },
              { id: 'left', label: 'Left border', icon: 'border_left' },
              { id: 'right', label: 'Right border', icon: 'border_right' },
              { id: 'none', label: 'No border', icon: 'border_clear' },
            ]}
            onDefault={(color) => api && setBorders(api, 'all', color)}
            onChoose={(choice, color) => api && setBorders(api, choice as BorderChoice, color)}
          />
        </RibbonGroup>

        <RibbonGroup label="Alignment">
          <ToolbarButton id="align-left" label="Align left" icon="format_align_left" pressed={state.align === 'left'} disabled={!ready} onClick={() => api && setAlignment(api, 'left')} />
          <ToolbarButton id="align-center" label="Center" icon="format_align_center" pressed={state.align === 'center'} disabled={!ready} onClick={() => api && setAlignment(api, 'center')} />
          <ToolbarButton id="align-right" label="Align right" icon="format_align_right" pressed={state.align === ('right' as HAlign)} disabled={!ready} onClick={() => api && setAlignment(api, 'right')} />
          <ToolbarButton id="align-top" label="Align top" icon="vertical_align_top" pressed={state.vAlign === 'top'} disabled={!ready} onClick={() => api && setVerticalAlignment(api, 'top')} />
          <ToolbarButton id="align-middle" label="Align middle" icon="vertical_align_center" pressed={state.vAlign === 'middle'} disabled={!ready} onClick={() => api && setVerticalAlignment(api, 'middle')} />
          <ToolbarButton id="align-bottom" label="Align bottom" icon="vertical_align_bottom" pressed={state.vAlign === ('bottom' as VAlign)} disabled={!ready} onClick={() => api && setVerticalAlignment(api, 'bottom')} />
          <ToolbarButton id="wrap-text" label="Wrap text" icon="wrap_text" pressed={state.isWrapped} disabled={!ready} onClick={() => api && toggleWrap(api, state.isWrapped)} />
          <ToolbarButton
            id="merge-cells"
            label={state.isMerged ? 'Unmerge cells' : 'Merge & Center'}
            icon={state.isMerged ? 'call_split' : 'cell_merge'}
            pressed={state.isMerged}
            disabled={!ready || (!state.isMerged && !state.isMultiCell)}
            onClick={() => api && toggleMerge(api, state.isMerged)}
          />
        </RibbonGroup>

        <RibbonGroup label="Number">
          <ToolbarSelect
            id="num-format"
            label="Number format"
            value={detectFormatKey(state.numberFormat)}
            options={NUMBER_FORMAT_OPTIONS}
            width={140}
            disabled={!ready}
            onChange={(v) => api && setNumberFormatByKey(api, v as NumberFormatKey)}
          />
          <ToolbarButton id="numfmt-currency" label="Currency" icon="attach_money" pressed={state.numberFormat === NUMBER_FORMATS.currency} disabled={!ready} onClick={() => api && setNumberFormat(api, NUMBER_FORMATS.currency)} />
          <ToolbarButton id="numfmt-percent" label="Percent" icon="percent" pressed={state.numberFormat === NUMBER_FORMATS.percent} disabled={!ready} onClick={() => api && setNumberFormat(api, NUMBER_FORMATS.percent)} />
        </RibbonGroup>

        <RibbonGroup label="Data">
          <ToolbarButton id="sort-asc" label="Sort ascending (A → Z)" icon="arrow_downward" disabled={!ready || !state.isMultiCell} onClick={() => api && sortRange(api, true)} />
          <ToolbarButton id="sort-desc" label="Sort descending (Z → A)" icon="arrow_upward" disabled={!ready || !state.isMultiCell} onClick={() => api && sortRange(api, false)} />
          <ToolbarButton id="filter-toggle" label="Toggle filter" icon="filter_alt" disabled={!ready || !state.isMultiCell} onClick={() => api && toggleFilter(api)} />
        </RibbonGroup>

        <RibbonGroup label="Insert">
          <ToolbarDropdown
            id="auto-sum"
            label="AutoSum"
            icon="functions"
            disabled={!ready}
            items={[
              { id: 'SUM', label: 'Sum', icon: 'functions' },
              { id: 'AVERAGE', label: 'Average', icon: 'bar_chart' },
              { id: 'COUNT', label: 'Count', icon: 'numbers' },
              { id: 'MAX', label: 'Max', icon: 'arrow_upward' },
              { id: 'MIN', label: 'Min', icon: 'arrow_downward' },
            ]}
            onDefault={() => api && applyAutoFunction(api, 'SUM')}
            onChoose={(id) => api && applyAutoFunction(api, id as 'SUM' | 'AVERAGE' | 'COUNT' | 'MIN' | 'MAX')}
          />
          <ToolbarDropdown
            id="format-as-table"
            label="Format as Table"
            icon="table_rows"
            disabled={!ready}
            items={[
              { id: 'plain', label: 'Plain (no style)', icon: 'grid_on' },
              ...TABLE_THEMES.map((t) => ({
                id: t.id,
                label: t.label,
                icon: 'table_chart',
              })),
            ]}
            onDefault={() => api && formatAsTable(api, 'table-default-0')}
            onChoose={(id) => {
              if (!api) return;
              if (id === 'plain') formatAsTable(api, undefined);
              else formatAsTable(api, id as TableThemeId);
            }}
          />
          <ToolbarButton
            id="tables-panel"
            label={ui.tablesPanelVisible ? 'Hide Tables panel' : 'Tables panel'}
            icon="view_sidebar"
            pressed={ui.tablesPanelVisible}
            onClick={ui.toggleTablesPanel}
          />
          <ToolbarButton id="insert-comment" label="Insert comment" icon="add_comment" disabled={!ready} onClick={() => api && insertComment(api)} />
          <ToolbarButton id="insert-hyperlink" label="Insert hyperlink (Ctrl+K)" icon="link" disabled={!ready} onClick={() => api && insertHyperlink(api)} />
        </RibbonGroup>
      </div>
    </nav>
  );
}
